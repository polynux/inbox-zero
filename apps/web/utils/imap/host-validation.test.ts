import { beforeEach, describe, expect, it, vi } from "vitest";
import { resolveHostAddresses } from "@inboxzero/network/pinned-dns";
import {
  assertImapHostAllowed,
  validateImapHosts,
} from "@/utils/imap/host-validation";

vi.mock("@inboxzero/network/pinned-dns", async (importOriginal) => {
  const actual =
    await importOriginal<typeof import("@inboxzero/network/pinned-dns")>();
  return {
    ...actual,
    resolveHostAddresses: vi.fn(actual.resolveHostAddresses),
  };
});

const blocked = { valid: false, error: "Host is not allowed" };

function resolvesTo(...addresses: string[]) {
  vi.mocked(resolveHostAddresses).mockResolvedValueOnce(
    addresses.map((address) => ({
      address,
      family: address.includes(":") ? 6 : 4,
    })),
  );
}

describe("validateImapHosts", () => {
  beforeEach(() => {
    vi.mocked(resolveHostAddresses).mockReset();
  });

  describe("private hosts blocked (default)", () => {
    it.each([
      ["localhost"],
      ["LOCALHOST."],
      ["mail.localhost"],
      ["bridge.local"],
      ["intranet"],
      ["metadata.google.internal"],
      ["127.0.0.1"],
      ["127.1"],
      ["2130706433"],
      ["0x7f000001"],
      ["0177.0.0.1"],
      ["10.0.0.5"],
      ["172.16.4.2"],
      ["192.168.1.10"],
      ["100.64.0.1"],
      ["169.254.169.254"],
      ["0.0.0.0"],
      ["[::1]"],
      ["[::ffff:127.0.0.1]"],
      ["::ffff:7f00:1"],
      ["[fe80::1]"],
      ["[fc00::1]"],
      ["[fd00:ec2::254]"],
    ])("rejects %s", async (host) => {
      expect(
        await validateImapHosts([host], { allowPrivateHosts: false }),
      ).toEqual(blocked);
    });

    it("rejects a public-looking name that resolves to a private address", async () => {
      resolvesTo("93.184.216.34", "10.0.0.5");

      expect(
        await validateImapHosts(["bridge.example.com"], {
          allowPrivateHosts: false,
        }),
      ).toEqual(blocked);
    });

    it("rejects a name that cannot be resolved", async () => {
      vi.mocked(resolveHostAddresses).mockRejectedValueOnce(
        Object.assign(new Error("not found"), { code: "ENOTFOUND" }),
      );

      expect(
        await validateImapHosts(["missing.example.com"], {
          allowPrivateHosts: false,
        }),
      ).toEqual({ valid: false, error: "Host could not be resolved" });
    });

    it("rejects when the SMTP host is private even if the IMAP host is public", async () => {
      resolvesTo("66.111.4.1");

      expect(
        await validateImapHosts(["imap.fastmail.com", "127.0.0.1"], {
          allowPrivateHosts: false,
        }),
      ).toEqual(blocked);
    });

    it("allows public names and addresses", async () => {
      resolvesTo("66.111.4.1");

      expect(
        await validateImapHosts(["imap.fastmail.com", "8.8.8.8"], {
          allowPrivateHosts: false,
        }),
      ).toEqual({ valid: true });
    });
  });

  describe("private hosts allowed (IMAP_ALLOW_PRIVATE_HOSTS)", () => {
    it.each([
      ["127.0.0.1"],
      ["::1"],
      ["[::1]"],
      ["[::ffff:127.0.0.1]"],
      ["10.0.0.5"],
      ["172.16.4.2"],
      ["192.168.1.10"],
      ["[fd12::1]"],
    ])("allows loopback and private %s", async (host) => {
      expect(
        await validateImapHosts([host], { allowPrivateHosts: true }),
      ).toEqual({ valid: true });
    });

    it("allows localhost", async () => {
      resolvesTo("127.0.0.1", "::1");

      expect(
        await validateImapHosts(["localhost"], { allowPrivateHosts: true }),
      ).toEqual({ valid: true });
    });

    it.each([
      ["169.254.169.254"],
      ["[::ffff:169.254.169.254]"],
      ["fd00:ec2::254"],
      ["metadata.google.internal"],
      ["169.254.1.1"],
      ["[fe80::1]"],
      ["0.0.0.0"],
      ["100.64.0.1"],
    ])("still rejects %s", async (host) => {
      expect(
        await validateImapHosts([host], { allowPrivateHosts: true }),
      ).toEqual(blocked);
    });

    it("still rejects a name that resolves to cloud metadata", async () => {
      resolvesTo("169.254.169.254");

      expect(
        await validateImapHosts(["bridge.lan"], { allowPrivateHosts: true }),
      ).toEqual(blocked);
    });
  });
});

describe("assertImapHostAllowed", () => {
  beforeEach(() => {
    vi.mocked(resolveHostAddresses).mockReset();
  });

  it("resolves when the stored host still validates", async () => {
    resolvesTo("66.111.4.1");

    await expect(
      assertImapHostAllowed("imap.fastmail.com", { allowPrivateHosts: false }),
    ).resolves.toBeUndefined();
  });

  it("throws when the stored host now resolves to a private address", async () => {
    resolvesTo("10.0.0.5");

    await expect(
      assertImapHostAllowed("imap.fastmail.com", { allowPrivateHosts: false }),
    ).rejects.toThrow("Host is not allowed");
  });
});
