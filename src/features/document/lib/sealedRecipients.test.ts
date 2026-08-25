import { describe, expect, it } from "vitest";

import {
  extraEncryptForUsernames,
  isSealedEncryptForDocument,
  preferDirectoryUsername,
  recipientDiff,
  recipientNamesEqual,
} from "./sealedRecipients";

describe("sealedRecipients", () => {
  it("treats canonical and abbreviated names as equal", () => {
    expect(recipientNamesEqual("cn=Ada/o=Acme", "Ada/Acme")).toBe(true);
    expect(recipientNamesEqual("CN=Ada/O=Acme", "cn=bob/o=acme")).toBe(false);
  });

  it("detects sealed documents from _encryptFor", () => {
    expect(isSealedEncryptForDocument({ title: "Note" })).toBe(false);
    expect(isSealedEncryptForDocument({ _encryptFor: {} })).toBe(true);
  });

  it("lists extra readers and skips the author, devices, and removed entries", () => {
    expect(
      extraEncryptForUsernames(
        {
          _encryptFor: {
            "cn=Maya Chen/o=Acme": { kind: "user" },
            "cn=Ada Lovelace/o=Acme": { kind: "user", label: "cn=Ada Lovelace/o=Acme" },
            "device#abc": { kind: "device" },
            "cn=Gone/o=Acme": { kind: "user", removedAt: 1 },
          },
        },
        "cn=Maya Chen/o=Acme",
      ),
    ).toEqual(["cn=Ada Lovelace/o=Acme"]);
  });

  it("diffs added and removed extra recipients", () => {
    expect(
      recipientDiff(["cn=Ada/o=Acme"], ["Ada/Acme", "cn=Bob/o=Acme"]),
    ).toEqual({
      added: ["cn=Bob/o=Acme"],
      removed: [],
    });
    expect(recipientDiff(["cn=Ada/o=Acme", "cn=Bob/o=Acme"], [])).toEqual({
      added: [],
      removed: ["cn=Ada/o=Acme", "cn=Bob/o=Acme"],
    });
  });

  it("prefers Directory casing over persist-key labels", () => {
    expect(
      preferDirectoryUsername("CN=maya chen/O=acme", [
        "cn=Ada Lovelace/o=Acme",
        "cn=Maya Chen/o=Acme",
      ]),
    ).toBe("cn=Maya Chen/o=Acme");
    expect(preferDirectoryUsername("cn=Ghost/o=Acme", ["cn=Ada Lovelace/o=Acme"])).toBe(
      "cn=Ghost/o=Acme",
    );
  });
});
