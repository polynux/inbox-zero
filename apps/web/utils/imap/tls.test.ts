import { describe, expect, it } from "vitest";
import { isLoopbackHost, shouldBypassTlsVerification } from "@/utils/imap/tls";

describe("isLoopbackHost", () => {
  it.each([
    ["127.0.0.1"],
    ["localhost"],
    ["::1"],
    ["[::1]"],
    ["LOCALHOST."],
  ])("treats %s as loopback", (host) => {
    expect(isLoopbackHost(host)).toBe(true);
  });

  it.each([
    ["192.168.1.10"],
    ["imap.fastmail.com"],
    ["10.0.0.5"],
  ])("does not treat %s as loopback", (host) => {
    expect(isLoopbackHost(host)).toBe(false);
  });
});

describe("shouldBypassTlsVerification", () => {
  it("bypasses verification for a loopback host when private hosts are allowed", () => {
    expect(
      shouldBypassTlsVerification("127.0.0.1", { allowPrivateHosts: true }),
    ).toBe(true);
  });

  it("never bypasses verification for a remote host", () => {
    expect(
      shouldBypassTlsVerification("imap.fastmail.com", {
        allowPrivateHosts: true,
      }),
    ).toBe(false);
  });

  it("never bypasses verification when private hosts are not allowed", () => {
    expect(
      shouldBypassTlsVerification("127.0.0.1", { allowPrivateHosts: false }),
    ).toBe(false);
  });
});
