import { afterEach, beforeEach, describe, expect, test } from "bun:test"
import { mkdir, rm } from "node:fs/promises"
import { join } from "node:path"
import { MemoryPlugin } from "../index"

const tempRoot = join(import.meta.dir, "..", ".tmp-tests")
let testDir = ""

const context = {
  sessionID: "test-session",
  messageID: "test-message",
  agent: "test-agent",
  abort: new AbortController().signal,
}

const loadTools = async () => {
  const plugin = await MemoryPlugin({ directory: testDir } as never)
  if (!plugin.tool) throw new Error("Plugin did not return tools")
  const recall = plugin.tool.memory_recall
  const remember = plugin.tool.memory_remember
  const exportMemories = plugin.tool.memory_export
  const importMemories = plugin.tool.memory_import
  const forget = plugin.tool.memory_forget
  const compact = plugin.tool.memory_compact
  const memoryContext = plugin.tool.memory_context
  if (!recall) throw new Error("Plugin did not return memory_recall")
  if (!remember) throw new Error("Plugin did not return memory_remember")
  if (!exportMemories) throw new Error("Plugin did not return memory_export")
  if (!importMemories) throw new Error("Plugin did not return memory_import")
  if (!forget) throw new Error("Plugin did not return memory_forget")
  if (!compact) throw new Error("Plugin did not return memory_compact")
  if (!memoryContext) throw new Error("Plugin did not return memory_context")
  return { recall, remember, exportMemories, importMemories, forget, compact, memoryContext }
}

beforeEach(async () => {
  testDir = join(tempRoot, crypto.randomUUID())
  await mkdir(join(testDir, ".opencode", "memory"), { recursive: true })
})

afterEach(async () => {
  await rm(testDir, { recursive: true, force: true })
})

describe("memory_recall", () => {
  test("returns the highest scoring query matches within the limit", async () => {
    await Bun.write(
      join(testDir, ".opencode", "memory", "2026-05-28.logfmt"),
      [
        'ts=2026-05-28T10:00:00.000Z type=context scope=api content="api only"',
        'ts=2026-05-28T10:01:00.000Z type=context scope=database content="api only"',
        'ts=2026-05-28T10:02:00.000Z type=decision scope=api content="api decision"',
      ].join("\n") + "\n",
    )

    const tools = await loadTools()
    const output = await tools.recall.execute({ query: "api", limit: 2 }, context)

    expect(output).toContain("[2026-05-28] decision/api: api decision")
    expect(output).toContain("[2026-05-28] context/api: api only")
    expect(output).not.toContain("context/database")
  })

  test("returns the latest chronological memories when no query is provided", async () => {
    await Bun.write(
      join(testDir, ".opencode", "memory", "2026-05-27.logfmt"),
      'ts=2026-05-27T10:00:00.000Z type=context scope=old content="old memory"\n',
    )
    await Bun.write(
      join(testDir, ".opencode", "memory", "2026-05-28.logfmt"),
      [
        'ts=2026-05-28T10:00:00.000Z type=context scope=first content="first new memory"',
        'ts=2026-05-28T11:00:00.000Z type=context scope=second content="second new memory"',
      ].join("\n") + "\n",
    )

    const tools = await loadTools()
    const output = await tools.recall.execute({ limit: 2 }, context)

    expect(output).toContain("[2026-05-28] context/first: first new memory")
    expect(output).toContain("[2026-05-28] context/second: second new memory")
    expect(output).not.toContain("old memory")
  })

  test("round-trips multiline content written by memory_remember", async () => {
    const tools = await loadTools()
    await tools.remember.execute(
      {
        type: "context",
        scope: "notes",
        content: "line one\nline two with \"quotes\" and \\ slash",
      },
      context,
    )

    const output = await tools.recall.execute({ scope: "notes", match: "exact" }, context)
    const raw = await Bun.file(join(testDir, ".opencode", "memory", new Date().toISOString().split("T")[0] + ".logfmt")).text()

    expect(output).toContain("line one\nline two with \"quotes\" and \\ slash")
    expect(raw).toContain('content="line one\\nline two with \\"quotes\\" and \\\\ slash"')
  })

  test("imports compatible logfmt records with escaped multiline content", async () => {
    const tools = await loadTools()
    await tools.importMemories.execute(
      {
        format: "logfmt",
        data: 'ts=2026-05-28T12:00:00.000Z type=context scope=imported content="first\\nsecond"',
      },
      context,
    )

    const output = await tools.recall.execute({ scope: "imported", match: "exact" }, context)
    const exported = await tools.exportMemories.execute({ format: "jsonl" }, context)

    expect(output).toContain("first\nsecond")
    expect(JSON.parse(exported).content).toBe("first\nsecond")
  })

  test("preserves raw backslashes from older compatible records", async () => {
    await Bun.write(
      join(testDir, ".opencode", "memory", "2026-05-28.logfmt"),
      'ts=2026-05-28T10:00:00.000Z type=context scope=paths content="C:\\tmp\\memory"\n',
    )

    const tools = await loadTools()
    const output = await tools.recall.execute({ scope: "paths", match: "exact" }, context)

    expect(output).toContain("C:\\tmp\\memory")
  })

  test("filters by tags, date range, and exact scope matching", async () => {
    await Bun.write(
      join(testDir, ".opencode", "memory", "2026-05-28.logfmt"),
      [
        'ts=2026-05-28T10:00:00.000Z type=context scope=api content="old api" tags=backend,stale',
        'ts=2026-05-28T11:00:00.000Z type=context scope=api-v2 content="new api v2" tags=backend,current',
        'ts=2026-05-28T12:00:00.000Z type=context scope=api content="new api" tags=backend,current',
      ].join("\n") + "\n",
    )

    const tools = await loadTools()
    const output = await tools.recall.execute(
      { scope: "api", match: "exact", tags: ["current"], since: "2026-05-28T11:30:00.000Z", until: "2026-05-28" },
      context,
    )

    expect(output).toContain("new api")
    expect(output).not.toContain("old api")
    expect(output).not.toContain("api-v2")
  })

  test("memory_forget with query deletes only the best matching memory", async () => {
    await Bun.write(
      join(testDir, ".opencode", "memory", "2026-05-28.logfmt"),
      [
        'ts=2026-05-28T10:00:00.000Z type=context scope=api content="keep postgres detail"',
        'ts=2026-05-28T11:00:00.000Z type=context scope=api content="delete redis detail"',
      ].join("\n") + "\n",
    )

    const tools = await loadTools()
    const deleted = await tools.forget.execute({ type: "context", scope: "api", reason: "test", query: "redis" }, context)
    const output = await tools.recall.execute({ scope: "api", match: "exact" }, context)

    expect(deleted).toContain("Deleted 1 context memory(s)")
    expect(output).toContain("keep postgres detail")
    expect(output).not.toContain("delete redis detail")
  })

  test("memory_export and memory_import round-trip json", async () => {
    const tools = await loadTools()
    await tools.remember.execute({ type: "pattern", scope: "tests", content: "use plugin interface", tags: ["testing"] }, context)

    const exported = await tools.exportMemories.execute({ format: "json" }, context)
    await rm(join(testDir, ".opencode", "memory"), { recursive: true, force: true })

    const imported = await tools.importMemories.execute({ format: "json", data: exported }, context)
    const output = await tools.recall.execute({ scope: "tests", match: "exact" }, context)

    expect(imported).toBe("Imported 1 memory(s)")
    expect(output).toContain("pattern/tests: use plugin interface [testing]")
  })

  test("memory_compact removes exact duplicate records", async () => {
    await Bun.write(
      join(testDir, ".opencode", "memory", "2026-05-28.logfmt"),
      [
        'ts=2026-05-28T10:00:00.000Z type=context scope=api content="duplicate"',
        'ts=2026-05-28T10:00:00.000Z type=context scope=api content="duplicate"',
      ].join("\n") + "\n",
    )

    const tools = await loadTools()
    const dryRun = await tools.compact.execute({ dryRun: true }, context)
    const compacted = await tools.compact.execute({}, context)
    const output = await tools.recall.execute({}, context)

    expect(dryRun).toContain("1 duplicate(s) removed")
    expect(compacted).toContain("1 duplicate(s) removed")
    expect(output).toContain("Found 1 memories")
  })

  test("memory_context returns a compact relevant memory pack", async () => {
    await Bun.write(
      join(testDir, ".opencode", "memory", "2026-05-28.logfmt"),
      [
        'ts=2026-05-28T10:00:00.000Z type=context scope=deploy/staging content="Use materialize-deployments.cjs for staging runtime restart" tags=staging,deploy',
        'ts=2026-05-28T11:00:00.000Z type=context scope=tests content="Run make staging-live-onboarding-e2e for staging onboarding" tags=staging,e2e',
        'ts=2026-05-28T12:00:00.000Z type=context scope=runtime/local content="Local Bifrost is available through host.docker.internal" tags=local',
      ].join("\n") + "\n",
    )

    const tools = await loadTools()
    const output = await tools.memoryContext.execute({ query: "staging deploy", limit: 2, maxChars: 220 }, context)

    expect(output).toContain("Relevant Memory:")
    expect(output).toContain("deploy/staging")
    expect(output).toContain("tests")
    expect(output).not.toContain("runtime/local")
  })

  test("automatic hooks are disabled by default", async () => {
    const plugin = await MemoryPlugin({ directory: testDir } as never)
    if (!plugin["chat.message"] || !plugin["experimental.chat.system.transform"] || !plugin.tool?.memory_recall) throw new Error("Plugin did not return hooks/tools")

    await Bun.write(
      join(testDir, ".opencode", "memory", "2026-05-28.logfmt"),
      'ts=2026-05-28T10:00:00.000Z type=context scope=deploy/staging content="Use materialize-deployments.cjs for staging runtime restart" tags=staging,deploy\n',
    )

    await plugin["chat.message"](
      { sessionID: "session-1", agent: "build", model: { providerID: "test", modelID: "test" } },
      {
        message: {} as never,
        parts: [{ type: "text", text: "remember that I prefer minimal diffs and how do I restart staging deployments?" }] as never,
      },
    )

    const system = { system: [] as string[] }
    await plugin["experimental.chat.system.transform"]({}, system)

    const output = await plugin.tool.memory_recall.execute({ scope: "user", match: "exact" }, context)

    expect(output).toContain("No matching memories")
    expect(system.system).toEqual([])
  })

  test("auto-save stores explicit remember requests when enabled", async () => {
    const plugin = await MemoryPlugin({ directory: testDir } as never, { autoSave: true })
    if (!plugin["chat.message"] || !plugin.tool?.memory_recall) throw new Error("Plugin did not return hooks/tools")

    await plugin["chat.message"](
      { sessionID: "session-1", agent: "build", model: { providerID: "test", modelID: "test" } },
      {
        message: {} as never,
        parts: [{ type: "text", text: "remember that I prefer minimal diffs" }] as never,
      },
    )

    const output = await plugin.tool.memory_recall.execute({ scope: "user", match: "exact" }, context)

    expect(output).toContain("preference/user: I prefer minimal diffs [auto]")
  })

  test("auto-load injects relevant memories into system context when enabled", async () => {
    await Bun.write(
      join(testDir, ".opencode", "memory", "2026-05-28.logfmt"),
      'ts=2026-05-28T10:00:00.000Z type=context scope=deploy/staging content="Use materialize-deployments.cjs for staging runtime restart" tags=staging,deploy\n',
    )
    const plugin = await MemoryPlugin({ directory: testDir } as never, { autoLoad: true })
    if (!plugin["chat.message"] || !plugin["experimental.chat.system.transform"]) throw new Error("Plugin did not return auto hooks")

    await plugin["chat.message"](
      { sessionID: "session-1", agent: "build", model: { providerID: "test", modelID: "test" } },
      {
        message: {} as never,
        parts: [{ type: "text", text: "how do I restart staging deployments?" }] as never,
      },
    )

    const output = { system: [] as string[] }
    await plugin["experimental.chat.system.transform"]({}, output)

    expect(output.system.join("\n")).toContain("Relevant Memory:")
    expect(output.system.join("\n")).toContain("deploy/staging")
  })
})
