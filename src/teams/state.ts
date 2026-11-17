/**
 * Ids of chats and channels this server created.
 *
 * `teams.send` refuses a target that was not created here unless the operator
 * listed it in `teams.allowed_targets`. That gate needs durable state, because
 * the process restarts between calls. File: `~/.config/teams-mcp/teams-state.json`
 * (0600, atomic, in-process serialised).
 */
import { chmod, mkdir, readFile, rename, writeFile } from "node:fs/promises";
import { randomUUID } from "node:crypto";
import { dirname } from "node:path";
import { config } from "../config.js";

export interface CreatedChat {
  surface: "chat";
  id: string;
  name?: string;
  created_at: string;
}

export interface CreatedChannel {
  surface: "channel";
  id: string;
  team_id: string;
  name?: string;
  created_at: string;
}

export type CreatedTarget = CreatedChat | CreatedChannel;

interface TeamsState {
  version: 1;
  targets: CreatedTarget[];
}


function isTarget(value: unknown): value is CreatedTarget {
  if (value === null || typeof value !== "object") return false;
  const row = value as Record<string, unknown>;
  if (typeof row.id !== "string" || typeof row.created_at !== "string") return false;
  if (row.surface === "chat") return true;
  return row.surface === "channel" && typeof row.team_id === "string";
}

/** Read the state file; a missing or corrupt document yields an empty ledger. */
export async function readTeamsState(path = config().teamsStatePath): Promise<CreatedTarget[]> {
  try {
    const parsed = JSON.parse(await readFile(path, "utf8")) as unknown;
    if (parsed === null || typeof parsed !== "object") return [];
    const targets = (parsed as Record<string, unknown>).targets;
    if (!Array.isArray(targets)) return [];
    return targets.filter(isTarget);
  } catch {
    return [];
  }
}

async function writeTeamsState(targets: CreatedTarget[], path: string): Promise<void> {
  await mkdir(dirname(path), { recursive: true, mode: 0o700 });
  const tmp = `${path}.${process.pid}.${randomUUID()}.tmp`;
  const doc: TeamsState = { version: 1, targets };
  await writeFile(tmp, `${JSON.stringify(doc, null, 2)}\n`, { mode: 0o600 });
  await rename(tmp, path);
  await chmod(path, 0o600).catch(() => undefined);
}

/** Serialise read-modify-write inside this process so parallel creates do not lose rows. */
let queue: Promise<unknown> = Promise.resolve();

export async function recordCreatedTarget(
  target: CreatedTarget,
  path = config().teamsStatePath,
): Promise<CreatedTarget[]> {
  const task = queue.then(async () => {
    const current = await readTeamsState(path);
    const next = [...current.filter((row) => !(row.surface === target.surface && row.id === target.id)), target];
    await writeTeamsState(next, path);
    return next;
  });
  queue = task.catch(() => undefined);
  return task;
}

/** Drop a target from the ledger (after a successful delete). Returns what remains. */
export async function removeCreatedTarget(
  query: { surface: "chat" | "channel"; id: string; teamId?: string },
  path = config().teamsStatePath,
): Promise<CreatedTarget[]> {
  const task = queue.then(async () => {
    const current = await readTeamsState(path);
    const next = current.filter((row) => {
      if (row.surface !== query.surface || row.id !== query.id) return true;
      return row.surface === "channel" && query.teamId !== undefined && row.team_id !== query.teamId;
    });
    if (next.length !== current.length) await writeTeamsState(next, path);
    return next;
  });
  queue = task.catch(() => undefined);
  return task;
}

/** True when this server created the exact surface (channel ids are team-scoped). */
export async function isCreatedTarget(
  query: { surface: "chat" | "channel"; id: string; teamId?: string },
  path = config().teamsStatePath,
): Promise<boolean> {
  const targets = await readTeamsState(path);
  return targets.some((row) => {
    if (row.surface !== query.surface || row.id !== query.id) return false;
    return row.surface === "chat" || row.team_id === query.teamId;
  });
}
