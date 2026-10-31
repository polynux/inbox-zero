import { useEffect, useMemo, useState } from "react";
import useSWR from "swr";
import type { LabelsResponse } from "@/app/api/labels/route";
import type { EmailLabel } from "@/providers/email-label-types";
import { useAccount } from "@/providers/EmailAccountProvider";
import { compareLabelsByName } from "@/utils/label/compare-labels";
import { supportsLabelsAction } from "@/utils/actions/mail";

export type UserLabel = {
  id: string;
  name: string;
  type: "user";
  labelListVisibility?: string;
  messageListVisibility?: string;
  color?: {
    textColor?: string | null;
    backgroundColor?: string | null;
  };
};

export type OutlookLabel = {
  id: string;
  name: string;
  type: "user";
  color?: string;
};

export type GenericLabel = UserLabel | OutlookLabel;

function isHidden(label: EmailLabel): boolean {
  return label.labelListVisibility === "labelHide";
}

function useLabelsResponse(explicitEmailAccountId?: string) {
  const {
    emailAccount,
    isLoading: isLoadingEmailAccount,
    providerRateLimit,
  } = useAccount();
  const key = explicitEmailAccountId
    ? ["/api/labels", explicitEmailAccountId]
    : "/api/labels";
  const swr = useSWR<LabelsResponse>(
    !isLoadingEmailAccount && emailAccount && !providerRateLimit ? key : null,
    { shouldRetryOnError: false },
  );

  return {
    ...swr,
    isLoading: isLoadingEmailAccount || swr.isLoading,
  };
}
/**
 * Whether "label" detaches a tag on this account (Gmail, Outlook, Proton
 * Bridge Labels/*) or moves the message (plain IMAP folders). The UI cannot
 * LIST from the browser, so this probes the provider once per account.
 */
export function useSupportsLabels(emailAccountId?: string) {
  const { provider } = useAccount();
  const [supportsLabels, setSupportsLabels] = useState<boolean | null>(null);
  useEffect(() => {
    if (!emailAccountId) {
      setSupportsLabels(null);
      return;
    }
    // Gmail/Outlook always detach labels; only IMAP varies (Bridge
    // Labels/* vs plain folders), so only it pays for a LIST probe.
    if (provider === "google" || provider === "microsoft") {
      setSupportsLabels(true);
      return;
    }
    if (provider !== "imap") {
      setSupportsLabels(false);
      return;
    }
    let cancelled = false;
    setSupportsLabels(null);
    supportsLabelsAction(emailAccountId, {})
      .then((result) => {
        if (!cancelled)
          setSupportsLabels(result?.data?.supportsLabels ?? false);
      })
      .catch(() => {
        if (!cancelled) setSupportsLabels(false);
      });
    return () => {
      cancelled = true;
    };
  }, [emailAccountId, provider]);
  return {
    supportsLabels: supportsLabels ?? false,
    isLoading: supportsLabels === null,
  };
}

export function useAllLabels() {
  const { data, isLoading, error, mutate } = useLabelsResponse();

  const userLabels = useMemo(() => {
    if (!data?.labels) return [];

    return data.labels
      .filter((label) => label.type === "user")
      .sort(compareLabelsByName);
  }, [data?.labels]);

  return {
    userLabels,
    data,
    isLoading,
    error,
    mutate,
  };
}

export function useLabels(emailAccountId?: string) {
  const { data, isLoading, error, mutate } = useLabelsResponse(emailAccountId);

  const userLabels: EmailLabel[] = useMemo(() => {
    if (!data?.labels) return [];

    return data.labels
      .filter((label) => label.type === "user")
      .map((label) => ({
        id: label.id || "",
        name: label.name || "",
        type: label.type || null,
        color: label.color,
        labelListVisibility: label.labelListVisibility,
        messageListVisibility: label.messageListVisibility,
      }))
      .sort(compareLabelsByName);
  }, [data?.labels]);

  return {
    userLabels,
    isLoading,
    error,
    mutate,
  };
}

export function useSplitLabels() {
  const { userLabels, isLoading, error, mutate } = useLabels();

  const { visibleLabels, hiddenLabels } = useMemo(() => {
    // Split labels into visible and hidden categories
    const visible: EmailLabel[] = [];
    const hidden: EmailLabel[] = [];

    userLabels.forEach((label) => {
      if (isHidden(label)) {
        hidden.push(label);
      } else {
        visible.push(label);
      }
    });

    return {
      visibleLabels: visible,
      hiddenLabels: hidden,
    };
  }, [userLabels]);

  return {
    visibleLabels,
    hiddenLabels,
    isLoading,
    error,
    mutate,
  };
}
