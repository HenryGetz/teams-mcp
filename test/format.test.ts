import { describe, expect, it } from "vitest";
import {
  errorText,
  htmlToText,
  isoTime,
  omitNullish,
  pick,
  previewFromBody,
  stableStringify,
  toJsonText,
  truncate,
} from "../src/format.js";

describe("truncate", () => {
  it("collapses whitespace and keeps short strings intact", () => {
    expect(truncate("  hello\n\tworld  ")).toBe("hello world");
  });

  it("cuts to the budget and marks the cut", () => {
    const out = truncate("x".repeat(50), 10);
    expect(out).toHaveLength(10);
    expect(out.endsWith("…")).toBe(true);
  });

  it("respects a custom budget and empty input", () => {
    expect(truncate("abc", 0)).toBe("");
    expect(truncate(null)).toBe("");
    expect(truncate(undefined)).toBe("");
  });

  it("defaults to 300 characters", () => {
    expect(truncate("y".repeat(1000))).toHaveLength(300);
  });
});

describe("omitNullish", () => {
  it("drops null and undefined members recursively", () => {
    const input = {
      a: 1,
      b: null,
      c: undefined,
      d: { e: null, f: "keep", g: [1, null, 2] },
      h: false,
      i: 0,
      j: "",
    };
    expect(omitNullish(input)).toEqual({
      a: 1,
      d: { f: "keep", g: [1, 2] },
      h: false,
      i: 0,
      j: "",
    });
  });

  it("keeps arrays that are not nullish", () => {
    expect(omitNullish({ list: [] })).toEqual({ list: [] });
  });
});

describe("stableStringify", () => {
  it("sorts keys so equal data serialises identically", () => {
    expect(stableStringify({ b: 1, a: { d: 2, c: 3 } })).toBe('{"a":{"c":3,"d":2},"b":1}');
  });

  it("orders arrays by position", () => {
    expect(stableStringify([{ b: 1, a: 2 }])).toBe('[{"a":2,"b":1}]');
  });
});

describe("toJsonText", () => {
  it("omits nulls and sorts keys", () => {
    expect(toJsonText({ z: null, a: 1 })).toBe('{"a":1}');
  });

  it("passes plain strings through without quoting", () => {
    expect(toJsonText("already text")).toBe("already text");
  });
});

describe("pick", () => {
  it("maps source fields onto short keys and drops missing values", () => {
    const row = { id: "1", subject: null, from: "a@b.c" };
    expect(pick(row, { id: "id", subject: "subject", from: (r) => r.from?.toUpperCase() })).toEqual({
      id: "1",
      from: "A@B.C",
    });
  });
});

describe("errorText", () => {
  it("renders one line with code, message and hint", () => {
    const text = errorText("unauthorized", "token\nexpired", "run auth login");
    expect(text).toBe("unauthorized: token expired (run auth login)");
    expect(text).not.toContain("\n");
  });
});

describe("isoTime", () => {
  it("normalises numbers, strings and dates", () => {
    const instant = Date.UTC(2026, 10, 16, 12, 0, 0);
    expect(isoTime(instant)).toBe("2026-11-16T12:00:00.000Z");
    expect(isoTime("2026-11-16T12:00:00Z")).toBe("2026-11-16T12:00:00.000Z");
    expect(isoTime(new Date(instant))).toBe("2026-11-16T12:00:00.000Z");
  });

  it("returns undefined for unusable input", () => {
    expect(isoTime(null)).toBeUndefined();
    expect(isoTime("not a date")).toBeUndefined();
  });
});

describe("previewFromBody", () => {
  it("strips markup and scripts before truncating", () => {
    const html = "<style>p{color:red}</style><p>Hello&nbsp;<b>there</b></p><script>evil()</script>";
    expect(previewFromBody(html)).toBe("Hello there");
  });
});

describe("htmlToText", () => {
  it("keeps paragraph breaks instead of flattening the mail", () => {
    expect(htmlToText("<p>first</p><p>second</p>")).toBe("first\nsecond");
    expect(htmlToText("line one<br>line two")).toBe("line one\nline two");
    expect(htmlToText("<div>a</div><div>b</div>")).toBe("a\nb");
  });

  it("decodes entities rather than returning them verbatim", () => {
    expect(htmlToText("<p>&quot;quoted&quot; &amp; &lt;angled&gt; &nbsp;done</p>")).toBe(
      '"quoted" & <angled> done',
    );
    expect(htmlToText("caf&#233; &mdash; 5 &#x2713;")).toBe("café — 5 ✓");
  });

  it("drops real markup but keeps text that only looks like a tag", () => {
    expect(htmlToText('<span style="color:red">red</span> text')).toBe("red text");
    expect(htmlToText("&lt;angle&gt; survives, <b>bold</b> does not")).toBe(
      "<angle> survives, bold does not",
    );
  });

  it("removes script/style blocks and comments", () => {
    expect(htmlToText("<style>p{}</style><!-- hidden --><p>visible</p><script>x()</script>")).toBe(
      "visible",
    );
  });

  it("renders list items on their own lines", () => {
    expect(htmlToText("<ul><li>one</li><li>two</li></ul>")).toBe("- one\n- two");
  });

  it("is idempotent on plain text and leaves unknown entities alone", () => {
    expect(htmlToText("plain\ntext with &unknown; entity")).toBe("plain\ntext with &unknown; entity");
  });

  it("feeds previewFromBody with a one-line summary", () => {
    expect(previewFromBody("<p>one</p><p>two</p>", 100)).toBe("one two");
  });
});
