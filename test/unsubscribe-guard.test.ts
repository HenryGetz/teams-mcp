/**
 * SSRF guard for one-click unsubscribe.
 *
 * `List-Unsubscribe` comes from an arbitrary sender, and `unsubscribe_apply`
 * POSTs to the URL in it. Before this suite the only check was a string prefix
 * test, so `https://169.254.169.254/...`, `https://localhost:8080/` and a
 * 302 into the private network all reached the network with no timeout.
 */
import { describe, expect, it } from "vitest";
import {
  assertSafeUnsubscribeTarget,
  isPrivateAddress,
  postOneClickUnsubscribe,
  setOneClickPoster,
  setOneClickResolver,
  UnsubscribeTargetError,
} from "../src/outlook/mail.js";

/** Poster that fails the test if the guard let a request through. */
function failPoster(calls: string[]) {
  setOneClickPoster(async (target) => {
    calls.push(target);
    return { status: 200 };
  });
}

async function refusal(target: string, resolve: () => Promise<string[]> = async () => ["93.184.216.34"]) {
  setOneClickResolver(resolve);
  const calls: string[] = [];
  failPoster(calls);
  let error: unknown;
  try {
    await postOneClickUnsubscribe(target);
  } catch (caught) {
    error = caught;
  }
  setOneClickPoster(undefined);
  setOneClickResolver(undefined);
  return { error, calls };
}

describe("unsubscribe target guard", () => {
  it("requires https", () => {
    expect(() => assertSafeUnsubscribeTarget("http://lists.example.com/u")).toThrow(UnsubscribeTargetError);
    expect(() => assertSafeUnsubscribeTarget("ftp://lists.example.com/u")).toThrow(/https/);
    expect(() => assertSafeUnsubscribeTarget("javascript:alert(1)")).toThrow(UnsubscribeTargetError);
    expect(() => assertSafeUnsubscribeTarget("not a url")).toThrow(/valid absolute URL/);
  });

  it("rejects embedded credentials", () => {
    expect(() => assertSafeUnsubscribeTarget("https://user:pw@lists.example.com/u")).toThrow(
      /must not embed credentials/,
    );
  });

  it("rejects loopback and internal hostnames before any DNS lookup", () => {
    for (const target of [
      "https://localhost/u",
      "https://foo.localhost/u",
      "https://printer.local/u",
      "https://metadata.google.internal/computeMetadata/v1",
      "https://svc.internal/u",
    ]) {
      expect(() => assertSafeUnsubscribeTarget(target)).toThrow(UnsubscribeTargetError);
    }
  });

  it("classifies private, loopback, link-local and mapped addresses", () => {
    for (const address of [
      "127.0.0.1",
      "0.0.0.0",
      "10.1.2.3",
      "172.16.5.5",
      "192.168.1.1",
      "169.254.169.254",
      "100.64.0.1",
      "::1",
      "::",
      "fe80::1",
      "fd00::1",
      "::ffff:127.0.0.1",
      "::ffff:10.0.0.1",
    ]) {
      expect(isPrivateAddress(address), address).toBe(true);
    }
    for (const address of ["93.184.216.34", "1.1.1.1", "2606:4700::1111"]) {
      expect(isPrivateAddress(address), address).toBe(false);
    }
  });

  it("refuses a public-looking host that resolves to a private address", async () => {
    const { error, calls } = await refusal("https://lists.example.com/u/1", async () => ["192.168.0.9"]);
    expect(error).toBeInstanceOf(UnsubscribeTargetError);
    expect((error as UnsubscribeTargetError).reason).toBe("private_address");
    expect(calls).toHaveLength(0);
  });

  it("refuses a literal private IP without resolving", async () => {
    const { error, calls } = await refusal("https://169.254.169.254/latest/meta-data/");
    expect((error as UnsubscribeTargetError).reason).toBe("private_address");
    expect(calls).toHaveLength(0);
  });

  it("refuses when DNS resolution fails, rather than guessing", async () => {
    const { error, calls } = await refusal("https://lists.example.com/u/1", async () => {
      throw new Error("ENOTFOUND");
    });
    expect((error as UnsubscribeTargetError).reason).toBe("dns_failed");
    expect(calls).toHaveLength(0);
  });

  it("posts once to a vetted public target", async () => {
    setOneClickResolver(async () => ["93.184.216.34"]);
    const posted: Array<{ target: string; body: string }> = [];
    setOneClickPoster(async (target, body) => {
      posted.push({ target, body });
      return { status: 204 };
    });
    const result = await postOneClickUnsubscribe("https://lists.example.com/u/1");
    setOneClickPoster(undefined);
    setOneClickResolver(undefined);
    expect(result.status).toBe(204);
    expect(posted).toEqual([
      { target: "https://lists.example.com/u/1", body: "List-Unsubscribe=One-Click" },
    ]);
  });

  it("does not follow a redirect out of the vetted host", async () => {
    setOneClickResolver(async () => ["93.184.216.34"]);
    setOneClickPoster(async () => ({ status: 302 }));
    let error: unknown;
    try {
      await postOneClickUnsubscribe("https://lists.example.com/u/1");
    } catch (caught) {
      error = caught;
    }
    setOneClickPoster(undefined);
    setOneClickResolver(undefined);
    expect((error as UnsubscribeTargetError).reason).toBe("redirect_not_followed");
  });
});
