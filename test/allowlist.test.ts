/**
 * Recipient-string and allowlist boundaries.
 *
 * The allowlist is the only thing standing between "an agent read an email from
 * an attacker" and "the agent mails an external address", so its matching rules
 * are pinned here: suffix matching must not accept `example.com.evil.com`, and a
 * display-name wrapper must not bypass (or spuriously trip) the check.
 */
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { resetConfigCache } from "../src/config.js";
import { disallowedRecipients } from "../src/tools/shared.js";
import { addrSpec } from "../src/tools/recipients.js";

let dir: string;
const saved = { ...process.env };

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "teams-mcp-allowlist-"));
  process.env.TEAMS_MCP_CONFIG_DIR = dir;
  resetConfigCache();
});

afterEach(async () => {
  process.env = { ...saved };
  resetConfigCache();
  await rm(dir, { recursive: true, force: true });
});

async function allow(list: string[]): Promise<void> {
  await writeFile(join(dir, "config.yaml"), `invite_allowlist: ${JSON.stringify(list)}\n`);
  resetConfigCache();
}

describe("invite_allowlist matching", () => {
  it("permits everything when the list is empty", async () => {
    expect(disallowedRecipients(["anyone@elsewhere.org"])).toEqual([]);
  });

  it("matches a domain suffix only on a label boundary", async () => {
    await allow(["@example.edu"]);
    expect(disallowedRecipients(["student@example.edu"])).toEqual([]);
    expect(disallowedRecipients(["Student@EXAMPLE.EDU"])).toEqual([]);
    for (const hostile of [
      "attacker@example.edu.evil.com",
      "attacker@notexample.edu",
      "attacker@example.edu@evil.com",
      "attacker@evil.com",
    ]) {
      expect(disallowedRecipients([hostile]), hostile).toEqual([hostile]);
    }
  });

  it("does not treat a subdomain as covered by the parent suffix", async () => {
    // Deliberate: a narrower rule is safer (see SECURITY.md, operator-only settings).
    await allow(["@example.edu"]);
    expect(disallowedRecipients(["member@example.edu"])).toEqual([]);
    expect(disallowedRecipients(["student@students.example.edu"])).toEqual([
      "student@students.example.edu",
    ]);
  });

  it("matches a full address exactly, case-insensitively", async () => {
    await allow(["student@example.edu"]);
    expect(disallowedRecipients(["Student@Example.edu"])).toEqual([]);
    expect(disallowedRecipients(["other@example.edu"])).toEqual(["other@example.edu"]);
    expect(disallowedRecipients(["student@example.edu.evil.com"])).toHaveLength(1);
  });

  it("checks a display-name wrapper against the address inside it", async () => {
    await allow(["@example.edu"]);
    // resolveRecipients normalises before the check; the allowlist sees the bare address.
    expect(addrSpec("Test Student <joel@example.edu>")).toBe(
      "joel@example.edu",
    );
    expect(disallowedRecipients([addrSpec("Joel <jd@example.edu>")])).toEqual([]);
    expect(disallowedRecipients([addrSpec("Joel <jd@evil.com>")])).toEqual(["jd@evil.com"]);
  });
});

describe("addrSpec", () => {
  it("extracts an angle-bracket address", () => {
    expect(addrSpec("Jane Doe <jane@example.com>")).toBe("jane@example.com");
    expect(addrSpec('<jane@example.com>')).toBe("jane@example.com");
  });

  it("unquotes a bare address and trims", () => {
    expect(addrSpec('  "jane@example.com" ')).toBe("jane@example.com");
    expect(addrSpec(" jane@example.com ")).toBe("jane@example.com");
  });

  it("leaves malformed multi-@ strings alone rather than guessing a prefix", () => {
    expect(addrSpec("a@b@c")).toBe("a@b@c");
    expect(addrSpec("not-an-address")).toBe("not-an-address");
  });
});
