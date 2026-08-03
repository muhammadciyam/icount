import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useRef, useState } from "react";
import inventory from "@/data/inventory.json";
import { supabase, supabaseEnabled } from "@/lib/supabase";

type TabId = "home" | "inventory" | "scan" | "reports" | "more";

type Item = {
  id: string;
  loc: string;
  name: string;
  qty: number;
  cost: number;
  price: number;
  // Outlet/branch this item's inventory belongs to. Every upload is tagged
  // with an outlet name, and outlets are tracked completely independently —
  // the same barcode uploaded for two outlets is two separate rows with
  // their own qty and their own counted qty (see keyOf() below).
  outlet: string;
  // "base" = part of the original bundled catalog, "custom" = added by staff.
  // Only present when items are loaded from Supabase — see supabase/schema.sql.
  source?: "base" | "custom";
  // Soft-delete flag: "Delete item" / "Delete all items" set this instead of
  // removing the row, so a deleted item can still be found (in the "Deleted"
  // filter) and restored later without losing its counted qty.
  deleted?: boolean;
};

const STORAGE_KEY = "stock-count-v1";
const INVENTORY_KEY = "stock-inventory-v1";
const AUTH_KEY = "stock-auth-v1";
const LAST_OUTLET_KEY = "stock-last-outlet-v1";
const SELECTED_OUTLET_KEY = "stock-selected-outlet-v1";
const DEFAULT_OUTLET = "Seven Mart";

// Sentinel for "don't filter by outlet" — never a real outlet name.
const ALL_OUTLETS = "__all_outlets__";

// Items/counts are keyed by outlet+id together (not just id) so the same
// barcode in two different outlets is tracked as two fully independent
// rows — separate qty, separate counted qty.
const SEP = "";
const keyOf = (outlet: string, id: string) => `${outlet}${SEP}${id}`;
const outletOfKey = (key: string) => key.slice(0, key.indexOf(SEP));

function getLastOutlet(): string {
  try {
    return localStorage.getItem(LAST_OUTLET_KEY) ?? DEFAULT_OUTLET;
  } catch {
    return DEFAULT_OUTLET;
  }
}

function setLastOutlet(outlet: string) {
  try {
    localStorage.setItem(LAST_OUTLET_KEY, outlet);
  } catch {
    /* ignore */
  }
}

// Local fallback admin login, used only when Supabase isn't configured (no
// database to store real accounts in). Once Supabase is set up, accounts
// live in the `staff_accounts` table — see supabase/schema.sql — and staff
// can be created/deleted/reset from the "Staff" panel in the app.
//
// This file is committed to git (and this repo is public), so this must
// never be a real account's password — it's a dead code path once Supabase
// is configured, but still ships in the client bundle either way.
const FALLBACK_ADMIN = { email: "siyante003@gmail.com", password: "HuujFC#pHgv5gdUxvkUA" };

type CountState = Record<string, number>;
type InventoryState = Record<string, Item>;
type StaffAccount = { id: string; email: string; name: string; role: "admin" | "staff" };

async function loginRequest(email: string, password: string): Promise<StaffAccount> {
  const normalizedEmail = email.trim().toLowerCase();
  if (supabaseEnabled && supabase) {
    const { data, error } = await supabase.rpc("staff_login", {
      p_email: normalizedEmail,
      p_password: password,
    });
    if (error) throw new Error(error.message);
    const row = (data as StaffAccount[] | null)?.[0];
    if (!row) throw new Error("Incorrect email or password.");
    return row;
  }
  if (normalizedEmail === FALLBACK_ADMIN.email && password === FALLBACK_ADMIN.password) {
    return { id: "local-admin", email: FALLBACK_ADMIN.email, name: "Admin", role: "admin" };
  }
  throw new Error("Incorrect email or password.");
}

const pick = (row: Record<string, unknown>, keys: string[]) => {
  for (const k of Object.keys(row)) {
    const norm = k.toLowerCase().replace(/[^a-z]/g, "");
    if (keys.includes(norm)) return row[k];
  }
  return undefined;
};

const num = (v: unknown) => {
  const n = Number(String(v ?? "").replace(/[^0-9.-]/g, ""));
  return Number.isFinite(n) ? n : 0;
};

async function parseSheet(file: File, outlet: string): Promise<Item[]> {
  const XLSX = await import("xlsx");
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0] ?? ""];
  if (!sheet) return [];

  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: "" });
  const out: Item[] = [];
  rows.forEach((row, i) => {
    const name = String(pick(row, ["name", "item", "itemname", "description", "product"]) ?? "").trim();
    const id = String(
      pick(row, ["id", "barcode", "code", "sku", "itemcode", "itemid"]) ?? "",
    ).trim();
    if (!name && !id) return;
    out.push({
      id: id || `X-${Date.now()}-${i}`,
      name: name || id,
      loc: String(pick(row, ["loc", "location", "shelf", "aisle", "rack"]) ?? "").trim(),
      qty: num(pick(row, ["qty", "quantity", "stock", "systemqty", "onhand", "balance"])),
      cost: num(pick(row, ["cost", "costprice", "purchaseprice", "buyprice"])),
      price: num(pick(row, ["price", "sellprice", "sellingprice", "retailprice", "unitprice"])),
      outlet,
    });
  });
  return out;
}

// Fetches every row of a table, paging past Supabase/PostgREST's default
// 1000-row response cap (a plain `.select("*")` silently truncates at 1000
// even when the table has far more rows — this is what previously made
// large uploads (thousands of items) appear incomplete in the app).
async function fetchAllRows<T>(
  client: NonNullable<typeof supabase>,
  table: string,
  columns: string,
  orderBy: string[],
): Promise<T[]> {
  const pageSize = 1000;
  const all: T[] = [];
  let from = 0;
  for (;;) {
    let q = client.from(table).select(columns).range(from, from + pageSize - 1);
    for (const col of orderBy) q = q.order(col, { ascending: true });
    const { data, error } = await q;
    if (error) throw error;
    const rows = (data ?? []) as T[];
    all.push(...rows);
    if (rows.length < pageSize) break;
    from += pageSize;
  }
  return all;
}


export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Stock Count — Seven Mart Inventory Counter" },
      {
        name: "description",
        content:
          "Search inventory items, enter counted quantity and tick them off. Track counting progress and variance against system stock.",
      },
      { property: "og:title", content: "Stock Count — Seven Mart Inventory Counter" },
      {
        property: "og:description",
        content: "Search inventory items, enter counted quantity and tick them off. Track counting progress and variance against system stock.",
      },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: Index,
});

function Index() {
  // Which bottom-nav tab is showing. Inventory is the default landing tab
  // since day-to-day counting is the app's main job; Home/Reports/More are
  // reachable via the bottom nav for overview/analysis/admin tasks without
  // cluttering the counting screen itself.
  const [activeTab, setActiveTab] = useState<TabId>("inventory");
  const [query, setQuery] = useState("");
  const [counts, setCounts] = useState<CountState>({});
  const [customItems, setCustomItems] = useState<InventoryState>({});
  const [filter, setFilter] = useState<"all" | "pending" | "counted" | "deleted">("all");
  const [showAdd, setShowAdd] = useState(false);
  const [addOutlet, setAddOutlet] = useState(DEFAULT_OUTLET);
  const [editingItem, setEditingItem] = useState<Item | null>(null);
  const [showStaffPanel, setShowStaffPanel] = useState(false);
  // Admin-only bulk selection ("Select items" toolbar button below) — lets
  // an admin cherry-pick specific items (within the current outlet/search/
  // filter view) and delete just those, instead of only "delete all".
  const [selectMode, setSelectMode] = useState(false);
  const [selectedKeys, setSelectedKeys] = useState<Set<string>>(new Set());
  const [loaded, setLoaded] = useState(false);
  const [account, setAccount] = useState<StaffAccount | null>(null);
  // In-memory only (never persisted) — lets admin actions in the staff
  // panel skip re-prompting for a password within the same page session.
  const [sessionPassword, setSessionPassword] = useState<string | null>(null);
  const [authChecked, setAuthChecked] = useState(false);
  // Which outlet's inventory is currently in view. ALL_OUTLETS shows every
  // outlet together (with an outlet badge per row); admin actions that
  // mutate data (reset counts / delete inventory) require a specific
  // outlet to be picked first, so a bulk action never silently hits every
  // outlet at once.
  const [selectedOutlet, setSelectedOutlet] = useState<string>(() => {
    try {
      return localStorage.getItem(SELECTED_OUTLET_KEY) ?? ALL_OUTLETS;
    } catch {
      return ALL_OUTLETS;
    }
  });

  // When Supabase is configured, the whole catalog (base + added) lives in
  // the `stock_items` table and is loaded into `customItems` below — the
  // bundled inventory.json is only used as an offline/no-database fallback.
  const items = useMemo<Item[]>(() => {
    if (supabaseEnabled) return Object.values(customItems);
    return [
      ...(inventory as Omit<Item, "outlet">[]).map((it) => ({
        ...it,
        outlet: DEFAULT_OUTLET,
        source: "base" as const,
      })),
      ...Object.values(customItems),
    ];
  }, [customItems]);

  // What the app treats as "the inventory" day to day — deleted items are
  // kept in `items` (and the database) so the "Deleted" filter can still
  // find them, but excluded everywhere else.
  const activeItems = useMemo(() => items.filter((i) => !i.deleted), [items]);

  // Every outlet name seen across all loaded items, for the outlet picker.
  const outlets = useMemo(
    () => Array.from(new Set(items.map((i) => i.outlet).filter(Boolean))).sort(),
    [items],
  );

  const selectOutlet = (outlet: string) => {
    setSelectedOutlet(outlet);
    try {
      localStorage.setItem(SELECTED_OUTLET_KEY, outlet);
    } catch {
      /* ignore */
    }
  };

  // Initial load: from Supabase (shared, cross-device) when configured,
  // otherwise from this browser's local storage only.
  useEffect(() => {
    async function load() {
      if (supabaseEnabled && supabase) {
        try {
          // .select() alone caps out at 1000 rows (Supabase/PostgREST's
          // default page size) — fetchAllRows pages past that so every
          // uploaded row is actually loaded, however many there are.
          const rows = await fetchAllRows<{ item_id: string; qty: number; outlet: string }>(
            supabase,
            "stock_counts",
            "item_id, qty, outlet",
            ["outlet", "item_id"],
          );
          setCounts(
            Object.fromEntries(
              rows.map((r) => [keyOf(r.outlet ?? DEFAULT_OUTLET, r.item_id), Number(r.qty)]),
            ),
          );
        } catch (err) {
          console.error("Failed to load shared counts", err);
        }
        try {
          const rows = await fetchAllRows<Record<string, unknown>>(
            supabase,
            "stock_items",
            "*",
            ["outlet", "id"],
          );
          setCustomItems(
            Object.fromEntries(
              rows.map((r) => {
                const outlet = (r["outlet"] as string | undefined) ?? DEFAULT_OUTLET;
                const item: Item = {
                  id: r["id"] as string,
                  name: r["name"] as string,
                  loc: (r["loc"] as string) ?? "",
                  qty: Number(r["qty"]),
                  cost: Number(r["cost"]),
                  price: Number(r["price"]),
                  outlet,
                  source: (r["source"] as "base" | "custom" | undefined) ?? "custom",
                  deleted: (r["deleted"] as boolean | undefined) ?? false,
                };
                return [keyOf(outlet, item.id), item];
              }),
            ),
          );
        } catch (err) {
          console.error("Failed to load shared items", err);
        }
      } else {
        try {
          const rawCounts = localStorage.getItem(STORAGE_KEY);
          if (rawCounts) setCounts(JSON.parse(rawCounts));
        } catch {
          /* ignore */
        }
        try {
          const rawItems = localStorage.getItem(INVENTORY_KEY);
          if (rawItems) setCustomItems(JSON.parse(rawItems));
        } catch {
          /* ignore */
        }
      }

      try {
        const raw = localStorage.getItem(AUTH_KEY);
        if (raw) setAccount(JSON.parse(raw) as StaffAccount);
      } catch {
        /* ignore */
      }
      setLoaded(true);
      setAuthChecked(true);
    }
    load();
  }, []);

  // Live sync: when Supabase is configured, push updates from other
  // devices/staff into this session as soon as they happen.
  useEffect(() => {
    if (!supabaseEnabled || !supabase) return;
    const client = supabase;

    const channel = client
      .channel("stock-sync")
      .on(
        "postgres_changes",
        { event: "*", schema: "public", table: "stock_counts" },
        (payload) => {
          setCounts((c) => {
            const next = { ...c };
            if (payload.eventType === "DELETE") {
              const old = payload.old;
              delete next[keyOf((old["outlet"] as string) ?? DEFAULT_OUTLET, old["item_id"] as string)];
            } else {
              const r = payload.new;
              next[keyOf((r["outlet"] as string) ?? DEFAULT_OUTLET, r["item_id"] as string)] = Number(
                r["qty"],
              );
            }
            return next;
          });
        },
      )
      .on(
        "postgres_changes",
        { event: "*", schema: "public", table: "stock_items" },
        (payload) => {
          setCustomItems((prev) => {
            const next = { ...prev };
            if (payload.eventType === "DELETE") {
              const old = payload.old;
              delete next[keyOf((old["outlet"] as string) ?? DEFAULT_OUTLET, old["id"] as string)];
            } else {
              const r = payload.new;
              const outlet = (r["outlet"] as string) ?? DEFAULT_OUTLET;
              next[keyOf(outlet, r["id"] as string)] = {
                id: r["id"] as string,
                name: r["name"] as string,
                loc: (r["loc"] as string) ?? "",
                qty: Number(r["qty"]),
                cost: Number(r["cost"]),
                price: Number(r["price"]),
                outlet,
                source: (r["source"] as "base" | "custom" | undefined) ?? "custom",
                deleted: (r["deleted"] as boolean | undefined) ?? false,
              };
            }
            return next;
          });
        },
      )
      .subscribe();

    return () => {
      client.removeChannel(channel);
    };
  }, []);

  // Local-storage fallback persistence (only used when Supabase isn't configured).
  useEffect(() => {
    if (!loaded || supabaseEnabled) return;
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(counts));
    } catch (err) {
      console.error("Failed to save counts", err);
      alert("Could not save your count — your device's storage may be full or unavailable.");
    }
  }, [counts, loaded]);

  useEffect(() => {
    if (!loaded || supabaseEnabled) return;
    try {
      localStorage.setItem(INVENTORY_KEY, JSON.stringify(customItems));
    } catch (err) {
      console.error("Failed to save imported items", err);
      alert("Could not save your imported items — your device's storage may be full or unavailable.");
    }
  }, [customItems, loaded]);

  const saveCount = async (outlet: string, itemId: string, qty: number) => {
    const k = keyOf(outlet, itemId);
    setCounts((c) => ({ ...c, [k]: qty }));
    if (supabaseEnabled && supabase) {
      const { error } = await supabase
        .from("stock_counts")
        .upsert(
          { outlet, item_id: itemId, qty, counted_by: account?.email ?? null, updated_at: new Date().toISOString() },
          { onConflict: "outlet,item_id" },
        );
      if (error) {
        console.error("Failed to save count", error);
        alert("Could not save your count to the shared database — check your connection and try again.");
      }
    }
  };

  const clearCount = async (outlet: string, itemId: string) => {
    const k = keyOf(outlet, itemId);
    setCounts((c) => {
      const n = { ...c };
      delete n[k];
      return n;
    });
    if (supabaseEnabled && supabase) {
      const { error } = await supabase.from("stock_counts").delete().eq("outlet", outlet).eq("item_id", itemId);
      if (error) {
        console.error("Failed to clear count", error);
        alert("Could not update the shared database — check your connection and try again.");
      }
    }
  };

  // Scoped to a single outlet on purpose — an admin must pick a specific
  // outlet before this is callable (see the "Reset count" button), so a
  // bulk reset never wipes every outlet's counts at once.
  const resetAllCounts = async (outlet: string) => {
    setCounts((c) => Object.fromEntries(Object.entries(c).filter(([k]) => outletOfKey(k) !== outlet)));
    if (supabaseEnabled && supabase) {
      const { error } = await supabase.from("stock_counts").delete().eq("outlet", outlet);
      if (error) {
        console.error("Failed to reset counts", error);
        alert("Could not reset the shared database — check your connection and try again.");
      }
    }
  };

  // Batched so a single large Excel upload (thousands of rows) doesn't hit
  // request-size/timeout limits in one giant upsert.
  const UPSERT_BATCH_SIZE = 500;

  const addItems = async (newItems: Item[]) => {
    setCustomItems((prev) => {
      const next = { ...prev };
      // Re-adding/re-importing an id that was previously deleted brings it
      // back into the active list, same as hitting "Restore".
      for (const it of newItems) next[keyOf(it.outlet, it.id)] = { ...it, deleted: false };
      return next;
    });
    if (supabaseEnabled && supabase) {
      for (let i = 0; i < newItems.length; i += UPSERT_BATCH_SIZE) {
        const batch = newItems.slice(i, i + UPSERT_BATCH_SIZE).map((it) => ({
          id: it.id,
          name: it.name,
          loc: it.loc,
          qty: it.qty,
          cost: it.cost,
          price: it.price,
          outlet: it.outlet,
          deleted: false,
          updated_at: new Date().toISOString(),
        }));
        const { error } = await supabase.from("stock_items").upsert(batch, { onConflict: "outlet,id" });
        if (error) {
          console.error("Failed to save items", error);
          alert("Could not save item(s) to the shared database — check your connection and try again.");
          break;
        }
      }
    }
  };

  // Edits an existing custom item's own details (name/location/system qty/
  // cost/price) — not to be confused with saveCount, which records what
  // staff *counted*. Restricted to custom rows, same as delete/restore, so
  // the bundled base catalog stays read-only. Outlet is fixed at creation
  // time and isn't editable here.
  const updateItem = async (updated: Item) => {
    const k = keyOf(updated.outlet, updated.id);
    setCustomItems((prev) => {
      const it = prev[k];
      if (!it || it.source === "base") return prev;
      return { ...prev, [k]: { ...it, ...updated } };
    });
    if (supabaseEnabled && supabase) {
      const { error } = await supabase
        .from("stock_items")
        .update({
          name: updated.name,
          loc: updated.loc,
          qty: updated.qty,
          cost: updated.cost,
          price: updated.price,
          updated_at: new Date().toISOString(),
        })
        .eq("outlet", updated.outlet)
        .eq("id", updated.id)
        .eq("source", "custom");
      if (error) {
        console.error("Failed to update item", error);
        alert("Could not save changes to the shared database — check your connection and try again.");
      }
    }
  };

  // Soft delete: items are never actually removed from the database, just
  // hidden from the normal list. Counts are left in place so restoring an
  // item (from the "Deleted" filter) brings back its last counted qty too.
  const deleteItem = async (item: Item) => {
    if (item.source === "base") return; // defense in depth — UI already hides this case
    const k = keyOf(item.outlet, item.id);
    setCustomItems((prev) => {
      const it = prev[k];
      if (!it) return prev;
      return { ...prev, [k]: { ...it, deleted: true } };
    });
    if (supabaseEnabled && supabase) {
      const { error } = await supabase
        .from("stock_items")
        .update({ deleted: true, updated_at: new Date().toISOString() })
        .eq("outlet", item.outlet)
        .eq("id", item.id)
        .eq("source", "custom");
      if (error) {
        console.error("Failed to delete item", error);
        alert("Could not delete that item from the shared database — check your connection and try again.");
      }
    }
  };

  const restoreItem = async (item: Item) => {
    const k = keyOf(item.outlet, item.id);
    setCustomItems((prev) => {
      const it = prev[k];
      if (!it) return prev;
      return { ...prev, [k]: { ...it, deleted: false } };
    });
    if (supabaseEnabled && supabase) {
      const { error } = await supabase
        .from("stock_items")
        .update({ deleted: false, updated_at: new Date().toISOString() })
        .eq("outlet", item.outlet)
        .eq("id", item.id);
      if (error) {
        console.error("Failed to restore item", error);
        alert("Could not restore that item in the shared database — check your connection and try again.");
      }
    }
  };

  // Soft-deletes everything added via "+ Add item" or Excel import for one
  // outlet — i.e. rows tagged source='custom' in that outlet. Rows tagged
  // source='base' (the original bundled catalog) are never touched by
  // this, and other outlets are untouched too: an admin must pick a
  // specific outlet before this is callable (see "Delete all items").
  // Deleted rows show up in the "Deleted" filter and can be restored
  // individually.
  const deleteAllCustomItems = async (outlet: string) => {
    const keys = Object.entries(customItems)
      .filter(([, it]) => it.outlet === outlet && it.source !== "base" && !it.deleted)
      .map(([k]) => k);
    setCustomItems((prev) => {
      const next = { ...prev };
      for (const k of keys) {
        const it = next[k];
        if (it) next[k] = { ...it, deleted: true };
      }
      return next;
    });
    if (supabaseEnabled && supabase) {
      const { error } = await supabase
        .from("stock_items")
        .update({ deleted: true, updated_at: new Date().toISOString() })
        .eq("outlet", outlet)
        .eq("source", "custom")
        .eq("deleted", false);
      if (error) {
        console.error("Failed to delete all items", error);
        alert("Could not delete all items from the shared database — check your connection and try again.");
      }
    }
  };

  // Soft-deletes a hand-picked set of items (the "Select items" bulk-select
  // flow below), as opposed to deleteAllCustomItems which wipes every
  // custom item in one outlet. Can span multiple outlets when "All
  // outlets" is the current view, so the writes are grouped per outlet.
  const deleteSelectedItems = async (selected: Item[]) => {
    const targets = selected.filter((it) => it.source !== "base" && !it.deleted);
    if (targets.length === 0) return;
    const keys = targets.map((it) => keyOf(it.outlet, it.id));
    setCustomItems((prev) => {
      const next = { ...prev };
      for (const k of keys) {
        const it = next[k];
        if (it) next[k] = { ...it, deleted: true };
      }
      return next;
    });
    if (supabaseEnabled && supabase) {
      const idsByOutlet = new Map<string, string[]>();
      for (const it of targets) {
        const arr = idsByOutlet.get(it.outlet) ?? [];
        arr.push(it.id);
        idsByOutlet.set(it.outlet, arr);
      }
      for (const [outlet, ids] of idsByOutlet) {
        const { error } = await supabase
          .from("stock_items")
          .update({ deleted: true, updated_at: new Date().toISOString() })
          .eq("outlet", outlet)
          .eq("source", "custom")
          .in("id", ids);
        if (error) {
          console.error("Failed to delete selected items", error);
          alert("Could not delete selected item(s) from the shared database — check your connection and try again.");
        }
      }
    }
  };

  // Items scoped to whichever outlet is currently selected (or every
  // outlet, when "All outlets" is picked) — used for the progress bar and
  // as the base list the search box and filter tabs narrow further.
  const outletScopedActiveItems = useMemo(
    () => (selectedOutlet === ALL_OUTLETS ? activeItems : activeItems.filter((i) => i.outlet === selectedOutlet)),
    [activeItems, selectedOutlet],
  );

  const results = useMemo(() => {
    const q = query.trim().toLowerCase();
    let list =
      filter === "deleted"
        ? items.filter((i) => i.deleted && (selectedOutlet === ALL_OUTLETS || i.outlet === selectedOutlet))
        : outletScopedActiveItems;
    if (q) {
      const terms = q.split(/\s+/);
      list = list.filter((i) => {
        const hay = `${i.id} ${i.name}`.toLowerCase();
        return terms.every((t) => hay.includes(t));
      });
    }
    if (filter === "pending") list = list.filter((i) => counts[keyOf(i.outlet, i.id)] === undefined);
    if (filter === "counted") list = list.filter((i) => counts[keyOf(i.outlet, i.id)] !== undefined);
    // No cap — every matching record is shown, however many there are.
    return list;
  }, [query, filter, counts, items, outletScopedActiveItems, selectedOutlet]);

  const done = outletScopedActiveItems.filter((i) => counts[keyOf(i.outlet, i.id)] !== undefined).length;
  const uncounted = outletScopedActiveItems.length - done;
  const pct = outletScopedActiveItems.length === 0 ? 0 : Math.round((done / outletScopedActiveItems.length) * 100);

  // Only currently-visible, deletable rows can be picked in bulk-select
  // mode — base catalog rows and (obviously) already-deleted rows can't.
  const selectableResults = useMemo(
    () => (filter === "deleted" ? [] : results.filter((i) => i.source !== "base")),
    [results, filter],
  );
  const allSelectableChosen =
    selectableResults.length > 0 && selectableResults.every((i) => selectedKeys.has(keyOf(i.outlet, i.id)));
  const toggleSelectAll = () => {
    setSelectedKeys((prev) => {
      if (allSelectableChosen) return new Set();
      return new Set(selectableResults.map((i) => keyOf(i.outlet, i.id)));
    });
  };
  const toggleSelectOne = (item: Item) => {
    const k = keyOf(item.outlet, item.id);
    setSelectedKeys((prev) => {
      const next = new Set(prev);
      if (next.has(k)) next.delete(k);
      else next.add(k);
      return next;
    });
  };

  // Reports tab: how far along each outlet is, computed independently of
  // whichever outlet the picker currently has selected.
  const perOutletProgress = useMemo(
    () =>
      outlets.map((o) => {
        const outletItems = activeItems.filter((i) => i.outlet === o);
        const doneCount = outletItems.filter((i) => counts[keyOf(o, i.id)] !== undefined).length;
        return { outlet: o, total: outletItems.length, done: doneCount };
      }),
    [outlets, activeItems, counts],
  );

  // Reports tab: items whose counted qty doesn't match system qty, scoped
  // to the same outlet the rest of the app is currently filtered to.
  const varianceList = useMemo(() => {
    const scope =
      selectedOutlet === ALL_OUTLETS ? activeItems : activeItems.filter((i) => i.outlet === selectedOutlet);
    return scope
      .map((item) => ({ item, counted: counts[keyOf(item.outlet, item.id)] }))
      .filter((x): x is { item: Item; counted: number } => x.counted !== undefined && x.counted !== x.item.qty);
  }, [activeItems, counts, selectedOutlet]);

  if (!authChecked) return null;

  if (!account) {
    return (
      <LoginScreen
        onSuccess={(acc, password) => {
          try {
            localStorage.setItem(AUTH_KEY, JSON.stringify(acc));
          } catch {
            /* ignore */
          }
          setAccount(acc);
          setSessionPassword(password);
        }}
      />
    );
  }

  const importExcelInput = (
    <input
      type="file"
      accept=".xlsx,.xls,.csv"
      className="hidden"
      onChange={async (e) => {
        const file = e.target.files?.[0];
        e.target.value = "";
        if (!file) return;
        const defaultOutlet = selectedOutlet !== ALL_OUTLETS ? selectedOutlet : getLastOutlet();
        const entered = prompt("Outlet name for this upload (e.g. branch/store name):", defaultOutlet);
        if (!entered || !entered.trim()) return;
        const outlet = entered.trim();
        try {
          const imported = await parseSheet(file, outlet);
          if (imported.length === 0) {
            alert("No rows found. Make sure the sheet has columns like Barcode/ID, Name, Qty.");
            return;
          }
          await addItems(imported);
          setLastOutlet(outlet);
          selectOutlet(outlet);
          alert(`Imported ${imported.length} items into "${outlet}".`);
        } catch (err) {
          alert("Could not read that file. Please upload a valid Excel or CSV file.");
          console.error(err);
        }
      }}
    />
  );

  const outletPicker = (
    <label className="flex items-center gap-2 text-xs text-muted-foreground">
      <span className="shrink-0 font-medium">Outlet</span>
      <select
        value={selectedOutlet}
        onChange={(e) => selectOutlet(e.target.value)}
        className="w-full rounded-lg border border-input bg-background px-3 py-1.5 text-sm text-foreground outline-none focus:ring-2 focus:ring-ring"
      >
        <option value={ALL_OUTLETS}>All outlets ({outlets.length})</option>
        {outlets.map((o) => (
          <option key={o} value={o}>
            {o}
          </option>
        ))}
      </select>
    </label>
  );

  return (
    <main className="min-h-screen bg-background pb-24">
      <header className="sticky top-0 z-10 border-b border-border bg-[#dae2e8] shadow-sm">
        <div className="mx-auto max-w-3xl px-4 py-2 sm:px-6">
          <div className="flex items-center gap-2.5">
            <div className="flex size-9 shrink-0 items-center justify-center rounded-xl bg-primary text-base font-bold text-primary-foreground shadow-sm">
              7M
            </div>
            <div className="min-w-0 flex-1">
              <h1 className="text-base font-semibold leading-tight tracking-tight text-foreground sm:text-lg">
                Stock Count
              </h1>
              <p className="truncate text-xs text-muted-foreground">
                {done} / {outletScopedActiveItems.length} counted · {pct}%
                {selectedOutlet !== ALL_OUTLETS && ` · ${selectedOutlet}`}
              </p>
            </div>
          </div>

          {activeTab === "inventory" && (
            <>
              <input
                autoFocus
                value={query}
                onChange={(e) => setQuery(e.target.value)}
                inputMode="search"
                placeholder="Search item name or barcode / ID…"
                className="mt-2 w-full rounded-xl border border-input bg-background px-4 py-2 text-base text-foreground outline-none placeholder:text-muted-foreground focus:ring-2 focus:ring-ring sm:py-1.5"
              />
              <div className="mt-1.5">{outletPicker}</div>
              <div className="mt-1.5 flex gap-1 rounded-lg border border-input p-1">
                {(["all", "pending", "counted", "deleted"] as const).map((f) => (
                  <button
                    key={f}
                    onClick={() => setFilter(f)}
                    className={`flex-1 rounded-md px-2 py-1.5 text-xs font-medium transition-colors ${
                      filter === f
                        ? "bg-primary text-primary-foreground"
                        : "text-muted-foreground hover:text-foreground"
                    }`}
                  >
                    {f === "all" ? "All" : f === "pending" ? "Not counted" : f === "counted" ? "Counted" : "Deleted"}
                  </button>
                ))}
              </div>

              {selectMode && (
                <div className="mt-1.5 flex items-center justify-between gap-2 rounded-lg border border-dashed border-input px-3 py-2 text-xs">
                  <label className="flex min-w-0 items-center gap-2 text-foreground">
                    <input
                      type="checkbox"
                      checked={allSelectableChosen}
                      onChange={toggleSelectAll}
                      disabled={selectableResults.length === 0}
                      className="size-4 shrink-0 accent-primary"
                    />
                    <span className="truncate">Select all ({selectableResults.length})</span>
                  </label>
                  <span className="shrink-0 text-muted-foreground">{selectedKeys.size} selected</span>
                  <button
                    type="button"
                    disabled={selectedKeys.size === 0}
                    onClick={() => {
                      const targets = items.filter((i) => selectedKeys.has(keyOf(i.outlet, i.id)));
                      if (targets.length === 0) return;
                      if (
                        confirm(
                          `Delete ${targets.length} selected item(s)? They can be restored from the "Deleted" filter.`,
                        )
                      ) {
                        void deleteSelectedItems(targets);
                        setSelectedKeys(new Set());
                      }
                    }}
                    className={`shrink-0 rounded-lg px-2.5 py-1.5 font-medium transition-colors ${
                      selectedKeys.size === 0
                        ? "cursor-not-allowed text-muted-foreground/50"
                        : "text-destructive hover:bg-destructive/10"
                    }`}
                  >
                    Delete selected
                  </button>
                </div>
              )}
            </>
          )}
        </div>
      </header>

      <div className="mx-auto max-w-3xl px-3 pb-6 sm:px-6">
        {activeTab === "home" && (
          <div className="space-y-4 py-4">
            <div className="grid grid-cols-3 gap-2">
              <div className="rounded-xl border border-border bg-card px-3 py-3">
                <p className="flex items-center gap-1.5 text-[11px] font-medium text-muted-foreground">
                  <span className="size-1.5 shrink-0 rounded-full bg-muted-foreground/50" aria-hidden="true" />
                  Total
                </p>
                <p className="mt-1 text-2xl font-semibold text-foreground">
                  {outletScopedActiveItems.length.toLocaleString()}
                </p>
              </div>
              <div className="rounded-xl border border-border bg-card px-3 py-3">
                <p className="flex items-center gap-1.5 text-[11px] font-medium text-muted-foreground">
                  <span className="size-1.5 shrink-0 rounded-full bg-primary" aria-hidden="true" />
                  Counted
                </p>
                <p className="mt-1 text-2xl font-semibold text-foreground">{done.toLocaleString()}</p>
              </div>
              <div className="rounded-xl border border-border bg-card px-3 py-3">
                <p className="flex items-center gap-1.5 text-[11px] font-medium text-muted-foreground">
                  <span className="size-1.5 shrink-0 rounded-full bg-destructive/60" aria-hidden="true" />
                  Uncounted
                </p>
                <p className="mt-1 text-2xl font-semibold text-foreground">{uncounted.toLocaleString()}</p>
              </div>
            </div>

            <div className="h-2 w-full overflow-hidden rounded-full bg-muted">
              <div
                className="h-full rounded-full bg-primary transition-all duration-300"
                style={{ width: `${pct}%` }}
              />
            </div>
            <p className="text-center text-xs text-muted-foreground">{pct}% counted</p>

            <div className="rounded-xl border border-border bg-card p-3">{outletPicker}</div>

            <button
              onClick={() => {
                setFilter("pending");
                setActiveTab("inventory");
              }}
              className="min-h-11 w-full rounded-xl bg-primary text-sm font-semibold text-primary-foreground shadow-sm transition-opacity hover:opacity-90"
            >
              Continue counting →
            </button>
          </div>
        )}

        {activeTab === "inventory" && (
          <>
            <p className="py-3 text-center text-xs text-muted-foreground">
              Showing all {results.length.toLocaleString()} item{results.length === 1 ? "" : "s"}
              {selectedOutlet === ALL_OUTLETS ? ` across ${outlets.length} outlet(s)` : ` in "${selectedOutlet}"`}
            </p>
            {results.length === 0 ? (
              <p className="py-16 text-center text-sm text-muted-foreground">No items found.</p>
            ) : (
              <ul className="space-y-2">
                {results.map((item) => {
                  const k = keyOf(item.outlet, item.id);
                  return (
                    <Row
                      key={k}
                      item={item}
                      counted={counts[k]}
                      isCustom={item.source !== "base"}
                      isAdmin={account.role === "admin"}
                      isDeleted={!!item.deleted}
                      showOutlet={selectedOutlet === ALL_OUTLETS}
                      selectMode={selectMode && item.source !== "base"}
                      selected={selectedKeys.has(k)}
                      onToggleSelect={() => toggleSelectOne(item)}
                      onSave={(v) => void saveCount(item.outlet, item.id, v)}
                      onClear={() => void clearCount(item.outlet, item.id)}
                      onEdit={() => setEditingItem(item)}
                      onDelete={() => void deleteItem(item)}
                      onRestore={() => void restoreItem(item)}
                    />
                  );
                })}
              </ul>
            )}
          </>
        )}

        {activeTab === "scan" && (
          <ScanTab
            onDetected={(text) => {
              setQuery(text);
              setFilter("all");
              selectOutlet(ALL_OUTLETS);
              setActiveTab("inventory");
            }}
          />
        )}

        {activeTab === "reports" && (
          <div className="space-y-5 py-4">
            <div>
              <h2 className="text-sm font-semibold text-foreground">Per-outlet progress</h2>
              <ul className="mt-2 space-y-2">
                {perOutletProgress.length === 0 && (
                  <p className="text-xs text-muted-foreground">No outlets yet.</p>
                )}
                {perOutletProgress.map((row) => {
                  const rowPct = row.total === 0 ? 0 : Math.round((row.done / row.total) * 100);
                  return (
                    <li key={row.outlet} className="rounded-lg border border-border bg-card px-3 py-2">
                      <div className="flex items-center justify-between gap-2 text-sm">
                        <span className="min-w-0 truncate font-medium text-foreground">{row.outlet}</span>
                        <span className="shrink-0 text-xs text-muted-foreground">
                          {row.done}/{row.total} · {rowPct}%
                        </span>
                      </div>
                      <div className="mt-1.5 h-1.5 w-full overflow-hidden rounded-full bg-muted">
                        <div
                          className="h-full rounded-full bg-primary transition-all duration-300"
                          style={{ width: `${rowPct}%` }}
                        />
                      </div>
                    </li>
                  );
                })}
              </ul>
            </div>

            <div>
              <h2 className="text-sm font-semibold text-foreground">
                Variances{selectedOutlet !== ALL_OUTLETS && ` — ${selectedOutlet}`}
              </h2>
              <p className="text-xs text-muted-foreground">
                Items whose counted quantity doesn't match the system quantity.
              </p>
              <ul className="mt-2 space-y-2">
                {varianceList.length === 0 && (
                  <p className="text-xs text-muted-foreground">No variances yet.</p>
                )}
                {varianceList.map(({ item, counted }) => {
                  const diff = counted - item.qty;
                  return (
                    <li
                      key={keyOf(item.outlet, item.id)}
                      className="rounded-lg border border-border bg-card px-3 py-2 text-sm"
                    >
                      <div className="flex items-center justify-between gap-2">
                        <span className="min-w-0 truncate font-medium text-foreground">{item.name}</span>
                        <span
                          className={`shrink-0 font-semibold ${diff > 0 ? "text-primary" : "text-destructive"}`}
                        >
                          {diff > 0 ? "+" : ""}
                          {diff}
                        </span>
                      </div>
                      <p className="text-xs text-muted-foreground">
                        {item.outlet} · #{item.id} · system {item.qty} · counted {counted}
                      </p>
                    </li>
                  );
                })}
              </ul>
            </div>
          </div>
        )}

        {activeTab === "more" && (
          <div className="space-y-2 py-4">
            <div className="rounded-xl border border-border bg-card p-3">{outletPicker}</div>

            <label className="flex min-h-11 cursor-pointer items-center gap-3 rounded-xl border border-border bg-card px-4 text-sm font-medium text-primary transition-colors hover:bg-primary/5">
              Import Excel
              {importExcelInput}
            </label>

            <button
              onClick={() => {
                let outlet = selectedOutlet;
                if (outlet === ALL_OUTLETS) {
                  const entered = prompt("Which outlet is this item for?", getLastOutlet());
                  if (!entered || !entered.trim()) return;
                  outlet = entered.trim();
                }
                setAddOutlet(outlet);
                setShowAdd(true);
              }}
              className="flex min-h-11 w-full items-center gap-3 rounded-xl border border-border bg-card px-4 text-left text-sm font-medium text-primary transition-colors hover:bg-primary/5"
            >
              + Add item
            </button>

            <button
              onClick={() => {
                if (account.role !== "admin") {
                  alert("Only an admin can select items to delete.");
                  return;
                }
                setSelectMode(true);
                setSelectedKeys(new Set());
                setActiveTab("inventory");
              }}
              title={account.role === "admin" ? undefined : "Admin only"}
              className={`flex min-h-11 w-full items-center gap-3 rounded-xl border border-border bg-card px-4 text-left text-sm font-medium transition-colors ${
                account.role === "admin"
                  ? "text-primary hover:bg-primary/5"
                  : "cursor-not-allowed text-muted-foreground/50"
              }`}
            >
              Select items…
            </button>

            <div className="pt-2">
              <p className="px-1 pb-1 text-[11px] font-medium uppercase tracking-wide text-muted-foreground">
                Destructive — outlet-scoped
              </p>
              <div className="space-y-2">
                <button
                  onClick={() => {
                    if (account.role !== "admin") {
                      alert("Only an admin can reset the count.");
                      return;
                    }
                    if (selectedOutlet === ALL_OUTLETS) {
                      alert("Select an outlet first — counts are reset one outlet at a time.");
                      return;
                    }
                    if (confirm(`Clear all counted items for "${selectedOutlet}"?`)) {
                      void resetAllCounts(selectedOutlet);
                    }
                  }}
                  title={account.role === "admin" ? undefined : "Admin only"}
                  className={`flex min-h-11 w-full items-center gap-3 rounded-xl border px-4 text-left text-sm font-medium transition-colors ${
                    account.role === "admin"
                      ? "border-destructive/30 text-destructive hover:bg-destructive/10"
                      : "cursor-not-allowed border-border text-muted-foreground/50"
                  }`}
                >
                  Reset count
                </button>

                <button
                  onClick={() => {
                    if (account.role !== "admin") {
                      alert("Only an admin can delete all items.");
                      return;
                    }
                    if (selectedOutlet === ALL_OUTLETS) {
                      alert("Select an outlet first — inventory is deleted one outlet at a time.");
                      return;
                    }
                    const activeCustomCount = Object.values(customItems).filter(
                      (it) => it.outlet === selectedOutlet && it.source !== "base" && !it.deleted,
                    ).length;
                    if (activeCustomCount === 0) {
                      alert("No added/imported items to delete for this outlet.");
                      return;
                    }
                    if (
                      confirm(
                        `Delete all ${activeCustomCount} added/imported item(s) in "${selectedOutlet}"? The base inventory list isn't affected. Deleted items can be restored from the "Deleted" filter.`,
                      )
                    ) {
                      void deleteAllCustomItems(selectedOutlet);
                    }
                  }}
                  title={account.role === "admin" ? undefined : "Admin only"}
                  className={`flex min-h-11 w-full items-center gap-3 rounded-xl border px-4 text-left text-sm font-medium transition-colors ${
                    account.role === "admin"
                      ? "border-destructive/30 text-destructive hover:bg-destructive/10"
                      : "cursor-not-allowed border-border text-muted-foreground/50"
                  }`}
                >
                  Delete all items
                </button>
              </div>
            </div>

            <div className="pt-2">
              <p className="px-1 pb-1 text-[11px] font-medium uppercase tracking-wide text-muted-foreground">
                Account
              </p>
              <div className="space-y-2">
                <div className="flex min-h-11 items-center justify-between rounded-xl border border-border bg-card px-4 text-sm">
                  <span className="text-muted-foreground">
                    Signed in as <span className="font-medium text-foreground">{account.name}</span>
                  </span>
                  {account.role === "admin" && (
                    <span className="rounded-full bg-primary/10 px-1.5 py-0.5 text-[10px] font-semibold text-primary">
                      Admin
                    </span>
                  )}
                </div>
                <button
                  onClick={() => setShowStaffPanel(true)}
                  className="flex min-h-11 w-full items-center gap-3 rounded-xl border border-border bg-card px-4 text-left text-sm font-medium text-foreground transition-colors hover:bg-muted"
                >
                  Staff accounts
                </button>
                <button
                  onClick={() => {
                    if (confirm("Log out?")) {
                      try {
                        localStorage.removeItem(AUTH_KEY);
                      } catch {
                        /* ignore */
                      }
                      setAccount(null);
                      setSessionPassword(null);
                    }
                  }}
                  className="flex min-h-11 w-full items-center gap-3 rounded-xl border border-border bg-card px-4 text-left text-sm font-medium text-foreground transition-colors hover:bg-muted"
                >
                  Log out
                </button>
              </div>
            </div>
          </div>
        )}
      </div>

      {showAdd && (
        <ItemFormModal
          title="Add new item"
          submitLabel="Save item"
          outlet={addOutlet}
          onClose={() => setShowAdd(false)}
          onSave={(item) => {
            void addItems([item]);
            setShowAdd(false);
          }}
        />
      )}

      {editingItem && (
        <ItemFormModal
          title="Edit item"
          submitLabel="Save changes"
          outlet={editingItem.outlet}
          initial={editingItem}
          onClose={() => setEditingItem(null)}
          onSave={(item) => {
            void updateItem(item);
            setEditingItem(null);
          }}
        />
      )}

      {showStaffPanel && (
        <StaffPanel
          account={account}
          sessionPassword={sessionPassword}
          onVerified={(pw) => setSessionPassword(pw)}
          onClose={() => setShowStaffPanel(false)}
        />
      )}

      <BottomNav active={activeTab} onChange={setActiveTab} />
    </main>
  );
}

const NAV_ITEMS: { id: TabId; label: string; icon: string }[] = [
  { id: "home", label: "Home", icon: "🏠" },
  { id: "inventory", label: "Inventory", icon: "📦" },
  { id: "scan", label: "Scan", icon: "📷" },
  { id: "reports", label: "Reports", icon: "📊" },
  { id: "more", label: "More", icon: "⚙️" },
];

// Fixed bottom tab bar — the app's primary navigation, kept within thumb
// reach for one-handed use. Each tab is a full-height, >=44px tap target
// (min-h-11) per mobile touch-target guidelines; the safe-area padding
// keeps it clear of the home-indicator area on notched phones.
function BottomNav({ active, onChange }: { active: TabId; onChange: (tab: TabId) => void }) {
  return (
    <nav
      className="fixed inset-x-0 bottom-0 z-20 border-t border-border bg-card shadow-[0_-1px_8px_rgba(0,0,0,0.06)]"
      style={{ paddingBottom: "env(safe-area-inset-bottom)" }}
    >
      <div className="mx-auto flex max-w-3xl items-stretch">
        {NAV_ITEMS.map((item) => {
          const isActive = active === item.id;
          return (
            <button
              key={item.id}
              type="button"
              onClick={() => onChange(item.id)}
              aria-current={isActive ? "page" : undefined}
              className={`flex min-h-11 flex-1 flex-col items-center justify-center gap-0.5 py-1.5 text-[10px] font-medium transition-colors ${
                isActive ? "text-primary" : "text-muted-foreground"
              }`}
            >
              <span className="text-lg leading-none" aria-hidden="true">
                {item.icon}
              </span>
              {item.label}
            </button>
          );
        })}
      </div>
    </nav>
  );
}

// Camera-based barcode scanner (Scan tab). Decodes continuously from the
// device's camera and reports the first successful read; the camera is
// stopped as soon as something is found (or the component unmounts) so it
// never keeps running in the background.
function ScanTab({ onDetected }: { onDetected: (text: string) => void }) {
  const videoRef = useRef<HTMLVideoElement>(null);
  const onDetectedRef = useRef(onDetected);
  onDetectedRef.current = onDetected;
  const [status, setStatus] = useState<"starting" | "scanning" | "error">("starting");
  const [error, setError] = useState<string | null>(null);
  const [restartToken, setRestartToken] = useState(0);

  useEffect(() => {
    let cancelled = false;
    let controls: { stop: () => void } | null = null;

    (async () => {
      try {
        const { BrowserMultiFormatReader } = await import("@zxing/browser");
        const reader = new BrowserMultiFormatReader();
        const c = await reader.decodeFromVideoDevice(undefined, videoRef.current ?? undefined, (result) => {
          if (result && !cancelled) {
            onDetectedRef.current(result.getText());
          }
        });
        if (cancelled) {
          c.stop();
          return;
        }
        controls = c;
        setStatus("scanning");
      } catch (err) {
        if (cancelled) return;
        console.error("Camera/scanner error", err);
        setError("Could not access the camera. Check camera permissions for this site and try again.");
        setStatus("error");
      }
    })();

    return () => {
      cancelled = true;
      controls?.stop();
    };
  }, [restartToken]);

  return (
    <div className="py-4">
      <div className="relative aspect-[3/4] w-full overflow-hidden rounded-2xl border border-border bg-black">
        <video ref={videoRef} muted playsInline autoPlay className="size-full object-cover" />
        {status === "scanning" && (
          <div className="pointer-events-none absolute inset-8 rounded-2xl border-2 border-white/70" />
        )}
        {status === "error" && (
          <div className="absolute inset-0 flex flex-col items-center justify-center gap-3 bg-black/80 p-6 text-center">
            <p className="text-sm text-white">{error}</p>
            <button
              type="button"
              onClick={() => {
                setStatus("starting");
                setError(null);
                setRestartToken((n) => n + 1);
              }}
              className="min-h-11 rounded-lg bg-white px-4 text-sm font-medium text-black"
            >
              Try again
            </button>
          </div>
        )}
      </div>
      <p className="mt-3 text-center text-sm text-muted-foreground">
        {status === "starting" && "Starting camera…"}
        {status === "scanning" && "Point the camera at a barcode."}
        {status === "error" && "Camera unavailable."}
      </p>
    </div>
  );
}

function Row({
  item,
  counted,
  isCustom,
  isAdmin,
  isDeleted,
  showOutlet,
  selectMode,
  selected,
  onToggleSelect,
  onSave,
  onClear,
  onEdit,
  onDelete,
  onRestore,
}: {
  item: Item;
  counted?: number | undefined;
  isCustom: boolean;
  isAdmin: boolean;
  isDeleted: boolean;
  showOutlet: boolean;
  // True only while the admin-only bulk "Select items" mode is on AND this
  // row is eligible (not a base-catalog row) — see the toolbar in Index().
  selectMode: boolean;
  selected: boolean;
  onToggleSelect: () => void;
  onSave: (v: number) => void;
  onClear: () => void;
  onEdit: () => void;
  onDelete: () => void;
  onRestore: () => void;
}) {
  const [open, setOpen] = useState(false);
  const [draft, setDraft] = useState("");
  const isDone = counted !== undefined;
  const diff = isDone ? counted - item.qty : 0;

  // Swipe left reveals Delete (right edge), swipe right reveals Edit (left
  // edge) — only for custom, non-deleted rows that aren't mid bulk-select.
  // The reveal is a two-step interaction (swipe, then tap the revealed
  // button, then confirm for delete) so it's no more accident-prone than
  // the equivalent buttons in the expanded row below; it's just reachable
  // one-handed without opening the row first.
  const swipeEnabled = isCustom && !isDeleted && !selectMode;
  const ACTION_WIDTH = 76;
  const [dragX, setDragX] = useState(0);
  const [swiping, setSwiping] = useState(false);
  const dragRef = useRef<{ startX: number; startY: number; base: number; axis: "x" | "y" | null } | null>(null);

  const closeSwipe = () => setDragX(0);

  const onSwipePointerDown = (e: React.PointerEvent<HTMLDivElement>) => {
    if (!swipeEnabled) return;
    dragRef.current = { startX: e.clientX, startY: e.clientY, base: dragX, axis: null };
  };
  const onSwipePointerMove = (e: React.PointerEvent<HTMLDivElement>) => {
    const st = dragRef.current;
    if (!st) return;
    const dx = e.clientX - st.startX;
    const dy = e.clientY - st.startY;
    if (st.axis === null) {
      if (Math.abs(dx) < 6 && Math.abs(dy) < 6) return;
      st.axis = Math.abs(dx) > Math.abs(dy) ? "x" : "y";
      if (st.axis === "x") {
        setSwiping(true);
        e.currentTarget.setPointerCapture(e.pointerId);
      }
    }
    if (st.axis !== "x") return;
    e.preventDefault();
    setDragX(Math.max(-ACTION_WIDTH, Math.min(ACTION_WIDTH, st.base + dx)));
  };
  const endSwipeDrag = () => {
    const st = dragRef.current;
    dragRef.current = null;
    setSwiping(false);
    if (!st || st.axis !== "x") return;
    setDragX((x) => (x > ACTION_WIDTH / 2 ? ACTION_WIDTH : x < -ACTION_WIDTH / 2 ? -ACTION_WIDTH : 0));
  };

  const borderClass = selected
    ? "border-primary"
    : isDeleted
      ? "border-border"
      : isDone
        ? "border-primary/30"
        : "border-border";
  const bgClass = selected
    ? "bg-primary/10"
    : isDeleted
      ? "bg-muted/30 opacity-70"
      : isDone
        ? "bg-primary/5"
        : "bg-card";

  return (
    <li className={`relative overflow-hidden rounded-xl border transition-colors ${borderClass}`}>
      {swipeEnabled && (
        <>
          <button
            type="button"
            tabIndex={-1}
            aria-hidden="true"
            onClick={() => {
              onEdit();
              closeSwipe();
            }}
            style={{ width: ACTION_WIDTH }}
            className="absolute inset-y-0 left-0 flex items-center justify-center bg-primary text-xs font-semibold text-primary-foreground"
          >
            Edit
          </button>
          <button
            type="button"
            tabIndex={-1}
            aria-hidden="true"
            onClick={() => {
              if (!isAdmin) {
                alert("Only an admin can delete items.");
                closeSwipe();
                return;
              }
              if (confirm(`Delete "${item.name}" from inventory?`)) onDelete();
              closeSwipe();
            }}
            style={{ width: ACTION_WIDTH }}
            className="absolute inset-y-0 right-0 flex items-center justify-center bg-destructive text-xs font-semibold text-destructive-foreground"
          >
            Delete
          </button>
        </>
      )}
      <div
        onPointerDown={onSwipePointerDown}
        onPointerMove={onSwipePointerMove}
        onPointerUp={endSwipeDrag}
        onPointerCancel={endSwipeDrag}
        style={{
          transform: `translateX(${dragX}px)`,
          transition: swiping ? "none" : "transform 0.2s ease",
          touchAction: swipeEnabled ? "pan-y" : undefined,
        }}
        className={`relative ${bgClass}`}
      >
        <div className="flex items-center gap-3 px-3 py-3">
          {selectMode ? (
            <input
              type="checkbox"
              aria-label={`Select ${item.name}`}
              checked={selected}
              onChange={onToggleSelect}
              className="size-5 shrink-0 accent-primary"
            />
          ) : (
            <button
              aria-label={isDone ? "Mark as not counted" : "Mark as counted"}
              onClick={() => {
                if (dragX !== 0) {
                  closeSwipe();
                  return;
                }
                if (isDeleted || !isDone) {
                  setOpen((o) => !o);
                  return;
                }
                onClear();
              }}
              disabled={isDeleted}
              className={`flex size-8 shrink-0 items-center justify-center rounded-full border text-sm transition-colors ${
                isDeleted
                  ? "border-input text-muted-foreground/40"
                  : isDone
                    ? "border-primary bg-primary text-primary-foreground"
                    : "border-input text-muted-foreground"
              }`}
            >
              {isDone ? "✓" : ""}
            </button>
          )}
          <button
            onClick={() => {
              if (dragX !== 0) {
                closeSwipe();
                return;
              }
              if (selectMode) onToggleSelect();
              else setOpen((o) => !o);
            }}
            className="min-w-0 flex-1 text-left"
          >
            <p className="flex items-center gap-1.5 truncate text-sm font-medium text-foreground">
              {showOutlet && (
                <span className="shrink-0 rounded-full bg-muted px-1.5 py-0.5 text-[10px] font-semibold text-muted-foreground">
                  {item.outlet}
                </span>
              )}
              <span className="truncate">{item.name}</span>
            </p>
            <p className="text-xs text-muted-foreground">
              #{item.id} · system {item.qty} · MVR {item.price.toFixed(2)}
              {isDone && (
                <span className={diff === 0 ? " font-medium text-primary" : " font-medium text-destructive"}>
                  {" "}
                  · counted {counted} ({diff > 0 ? "+" : ""}
                  {diff})
                </span>
              )}
            </p>
          </button>
        </div>

        {open && (
        <div className="space-y-3 border-t border-border bg-muted/40 px-3 py-3">
          <div className="grid grid-cols-2 gap-2 text-xs text-muted-foreground">
            <span>Outlet: <span className="text-foreground">{item.outlet}</span></span>
            <span>Location: <span className="text-foreground">{item.loc}</span></span>
            <span>System qty: <span className="text-foreground">{item.qty}</span></span>
            <span>Cost: <span className="text-foreground">{item.cost.toFixed(3)}</span></span>
            <span>Sale price: <span className="text-foreground">{item.price.toFixed(2)}</span></span>
          </div>
          {isDeleted ? (
            <>
              <p className="text-xs text-muted-foreground">
                Deleted — still in the database with its last counted qty. Restore it to count or edit it again.
              </p>
              {isCustom && (
                <button
                  type="button"
                  title={isAdmin ? undefined : "Admin only"}
                  onClick={() => {
                    if (!isAdmin) {
                      alert("Only an admin can restore items.");
                      return;
                    }
                    onRestore();
                  }}
                  className={`w-full rounded-lg border px-3 py-2 text-sm font-medium transition-colors ${
                    isAdmin
                      ? "border-primary/40 text-primary hover:bg-primary/10"
                      : "cursor-not-allowed border-input text-muted-foreground/50"
                  }`}
                >
                  Restore item
                </button>
              )}
            </>
          ) : (
            <>
              <form
                className="flex gap-2"
                onSubmit={(e) => {
                  e.preventDefault();
                  const v = Number(draft);
                  if (draft.trim() === "" || Number.isNaN(v)) return;
                  onSave(v);
                  setDraft("");
                  setOpen(false);
                }}
              >
                <input
                  type="number"
                  step="any"
                  value={draft}
                  onChange={(e) => setDraft(e.target.value)}
                  placeholder="Counted qty"
                  className="flex-1 rounded-lg border border-input bg-background px-3 py-2 text-base text-foreground outline-none focus:ring-2 focus:ring-ring"
                />
                <button
                  type="submit"
                  className="rounded-lg bg-primary px-4 py-2 text-sm font-medium text-primary-foreground"
                >
                  Save
                </button>
                <button
                  type="button"
                  onClick={() => {
                    onSave(item.qty);
                    setOpen(false);
                  }}
                  className="rounded-lg border border-input px-3 py-2 text-sm text-foreground"
                >
                  Match
                </button>
              </form>
              {isCustom && (
                <div className="flex gap-2">
                  <button
                    type="button"
                    onClick={onEdit}
                    className="flex-1 rounded-lg border border-input px-3 py-2 text-sm font-medium text-foreground transition-colors hover:bg-muted"
                  >
                    Edit item
                  </button>
                  <button
                    type="button"
                    title={isAdmin ? undefined : "Admin only"}
                    onClick={() => {
                      if (!isAdmin) {
                        alert("Only an admin can delete items.");
                        return;
                      }
                      if (confirm(`Delete "${item.name}" from inventory?`)) onDelete();
                    }}
                    className={`flex-1 rounded-lg border px-3 py-2 text-sm font-medium transition-colors ${
                      isAdmin
                        ? "border-destructive/40 text-destructive hover:bg-destructive/10"
                        : "cursor-not-allowed border-input text-muted-foreground/50"
                    }`}
                  >
                    Delete item
                  </button>
                </div>
              )}
            </>
          )}
        </div>
        )}
      </div>
    </li>
  );
}

function ItemFormModal({
  title,
  submitLabel,
  outlet,
  initial,
  onClose,
  onSave,
}: {
  title: string;
  submitLabel: string;
  // Outlet this item belongs to. Fixed for the lifetime of the modal — set
  // once (from the outlet picker, or a prompt) before it opens, and not
  // editable here, since items are keyed by outlet+id together.
  outlet: string;
  initial?: Item;
  onClose: () => void;
  onSave: (item: Item) => void;
}) {
  const [name, setName] = useState(initial?.name ?? "");
  const [loc, setLoc] = useState(initial?.loc ?? "");
  const [qty, setQty] = useState(initial ? String(initial.qty) : "");
  const [cost, setCost] = useState(initial ? String(initial.cost) : "");
  const [price, setPrice] = useState(initial ? String(initial.price) : "");

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const n = name.trim();
    const qNum = Number(qty);
    const cNum = Number(cost);
    const pNum = Number(price);
    if (!n || Number.isNaN(qNum) || Number.isNaN(cNum) || Number.isNaN(pNum)) return;
    onSave({
      ...initial,
      id: initial?.id ?? `custom-${Date.now()}`,
      outlet,
      name: n,
      loc: loc.trim() || "N/A",
      qty: qNum,
      cost: cNum,
      price: pNum,
    });
  };

  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-black/40 p-4 sm:items-center">
      <div className="w-full max-w-md rounded-2xl border border-border bg-card p-5 shadow-lg">
        <div className="mb-4 flex items-center justify-between">
          <div>
            <h2 className="text-lg font-semibold text-foreground">{title}</h2>
            <p className="text-xs text-muted-foreground">Outlet: {outlet}</p>
          </div>
          <button
            onClick={onClose}
            className="rounded-md p-1 text-muted-foreground hover:bg-muted"
            aria-label="Close"
          >
            ×
          </button>
        </div>
        <form onSubmit={handleSubmit} className="space-y-3">
          <input
            value={name}
            onChange={(e) => setName(e.target.value)}
            placeholder="Item name *"
            maxLength={120}
            required
            className="w-full rounded-lg border border-input bg-background px-3 py-2 text-base text-foreground outline-none focus:ring-2 focus:ring-ring"
          />
          <input
            value={loc}
            onChange={(e) => setLoc(e.target.value)}
            placeholder="Location"
            maxLength={60}
            className="w-full rounded-lg border border-input bg-background px-3 py-2 text-base text-foreground outline-none focus:ring-2 focus:ring-ring"
          />
          <div className="grid grid-cols-3 gap-2">
            <input
              type="number"
              step="any"
              value={qty}
              onChange={(e) => setQty(e.target.value)}
              placeholder="Qty *"
              required
              className="w-full rounded-lg border border-input bg-background px-3 py-2 text-base text-foreground outline-none focus:ring-2 focus:ring-ring"
            />
            <input
              type="number"
              step="any"
              value={cost}
              onChange={(e) => setCost(e.target.value)}
              placeholder="Cost *"
              required
              className="w-full rounded-lg border border-input bg-background px-3 py-2 text-base text-foreground outline-none focus:ring-2 focus:ring-ring"
            />
            <input
              type="number"
              step="any"
              value={price}
              onChange={(e) => setPrice(e.target.value)}
              placeholder="Price *"
              required
              className="w-full rounded-lg border border-input bg-background px-3 py-2 text-base text-foreground outline-none focus:ring-2 focus:ring-ring"
            />
          </div>
          <button
            type="submit"
            className="w-full rounded-lg bg-primary py-2.5 text-sm font-medium text-primary-foreground"
          >
            {submitLabel}
          </button>
        </form>
      </div>
    </div>
  );
}

function LoginScreen({ onSuccess }: { onSuccess: (account: StaffAccount, password: string) => void }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setBusy(true);
    setError(null);
    try {
      const account = await loginRequest(email, password);
      onSuccess(account, password);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Incorrect email or password.");
    } finally {
      setBusy(false);
    }
  };

  return (
    <main className="flex min-h-screen items-center justify-center bg-background px-4">
      <div className="w-full max-w-sm rounded-2xl border border-border bg-card p-6 shadow-lg">
        <div className="flex flex-col items-center text-center">
          <div className="flex size-14 items-center justify-center rounded-2xl bg-primary text-xl font-bold text-primary-foreground shadow-sm">
            7M
          </div>
          <h1 className="mt-4 text-xl font-semibold tracking-tight text-foreground">Stock Count</h1>
          <p className="mt-1 text-sm text-muted-foreground">Enter your email and password to continue.</p>
        </div>
        <form onSubmit={handleSubmit} className="mt-6 space-y-3">
          <input
            autoFocus
            type="email"
            value={email}
            onChange={(e) => {
              setEmail(e.target.value);
              setError(null);
            }}
            placeholder="Email"
            className={`w-full rounded-xl border bg-background px-4 py-3 text-base text-foreground outline-none placeholder:text-muted-foreground focus:ring-2 focus:ring-ring ${
              error ? "border-destructive" : "border-input"
            }`}
          />
          <input
            type="password"
            value={password}
            onChange={(e) => {
              setPassword(e.target.value);
              setError(null);
            }}
            placeholder="Password"
            className={`w-full rounded-xl border bg-background px-4 py-3 text-base text-foreground outline-none placeholder:text-muted-foreground focus:ring-2 focus:ring-ring ${
              error ? "border-destructive" : "border-input"
            }`}
          />
          {error && <p className="text-sm text-destructive">{error}</p>}
          <button
            type="submit"
            disabled={busy}
            className="w-full rounded-xl bg-primary py-3 text-sm font-semibold text-primary-foreground shadow-sm transition-opacity hover:opacity-90 disabled:opacity-60"
          >
            {busy ? "Logging in…" : "Log in"}
          </button>
        </form>
      </div>
    </main>
  );
}

function StaffPanel({
  account,
  sessionPassword,
  onVerified,
  onClose,
}: {
  account: StaffAccount;
  sessionPassword: string | null;
  onVerified: (password: string) => void;
  onClose: () => void;
}) {
  const isAdmin = account.role === "admin";
  const canManage = isAdmin && !!sessionPassword;

  const [staff, setStaff] = useState<StaffAccount[]>([]);
  const [loading, setLoading] = useState(true);
  const [confirmPw, setConfirmPw] = useState("");
  const [confirmError, setConfirmError] = useState<string | null>(null);
  const [confirmBusy, setConfirmBusy] = useState(false);
  const [showCreate, setShowCreate] = useState(false);

  const refresh = async () => {
    if (!supabaseEnabled || !supabase) return;
    const { data, error } = await supabase.rpc("staff_list");
    if (!error) setStaff((data ?? []) as StaffAccount[]);
  };

  useEffect(() => {
    void refresh().finally(() => setLoading(false));
  }, []);

  const handleConfirm = async (e: React.FormEvent) => {
    e.preventDefault();
    setConfirmBusy(true);
    setConfirmError(null);
    try {
      await loginRequest(account.email, confirmPw);
      onVerified(confirmPw);
    } catch {
      setConfirmError("Incorrect password.");
    } finally {
      setConfirmBusy(false);
    }
  };

  const handleDelete = async (target: StaffAccount) => {
    if (!supabase || !sessionPassword) return;
    if (target.id === account.id) {
      alert("You can't delete your own account.");
      return;
    }
    if (!confirm(`Delete ${target.name} (${target.email})?`)) return;
    const { error } = await supabase.rpc("staff_delete_user", {
      p_admin_email: account.email,
      p_admin_password: sessionPassword,
      p_target_id: target.id,
    });
    if (error) alert(error.message || "Could not delete that account.");
    else void refresh();
  };

  const handleReset = async (target: StaffAccount) => {
    if (!supabase || !sessionPassword) return;
    const next = prompt(`New password for ${target.name}:`);
    if (!next) return;
    const { error } = await supabase.rpc("staff_reset_password", {
      p_admin_email: account.email,
      p_admin_password: sessionPassword,
      p_target_id: target.id,
      p_new_password: next,
    });
    if (error) alert(error.message || "Could not reset that password.");
    else alert(`Password reset for ${target.name}.`);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-black/40 p-4 sm:items-center">
      <div className="max-h-[85vh] w-full max-w-md overflow-y-auto rounded-2xl border border-border bg-card p-5 shadow-lg">
        <div className="mb-4 flex items-center justify-between">
          <h2 className="text-lg font-semibold text-foreground">Staff accounts</h2>
          <button onClick={onClose} className="rounded-md p-1 text-muted-foreground hover:bg-muted" aria-label="Close">
            ×
          </button>
        </div>

        {!supabaseEnabled ? (
          <p className="text-sm text-muted-foreground">
            Connect Supabase (see .env) to manage staff accounts across devices.
          </p>
        ) : (
          <>
            {isAdmin && !canManage && (
              <form onSubmit={handleConfirm} className="mb-4 space-y-2 rounded-xl border border-border bg-muted/40 p-3">
                <p className="text-xs text-muted-foreground">
                  Re-enter your password to create, delete, or reset staff.
                </p>
                <input
                  type="password"
                  autoFocus
                  value={confirmPw}
                  onChange={(e) => {
                    setConfirmPw(e.target.value);
                    setConfirmError(null);
                  }}
                  placeholder="Your password"
                  className="w-full rounded-lg border border-input bg-background px-3 py-2 text-sm text-foreground outline-none focus:ring-2 focus:ring-ring"
                />
                {confirmError && <p className="text-xs text-destructive">{confirmError}</p>}
                <button
                  type="submit"
                  disabled={confirmBusy}
                  className="rounded-lg bg-primary px-3 py-1.5 text-xs font-medium text-primary-foreground disabled:opacity-60"
                >
                  {confirmBusy ? "Checking…" : "Confirm"}
                </button>
              </form>
            )}

            {loading ? (
              <p className="text-sm text-muted-foreground">Loading…</p>
            ) : (
              <ul className="space-y-2">
                {staff.map((s) => (
                  <li
                    key={s.id}
                    className="flex items-center justify-between gap-2 rounded-lg border border-border px-3 py-2"
                  >
                    <div className="min-w-0">
                      <p className="truncate text-sm font-medium text-foreground">
                        {s.name}
                        {s.role === "admin" && (
                          <span className="ml-1.5 rounded-full bg-primary/10 px-1.5 py-0.5 text-[10px] font-semibold text-primary">
                            Admin
                          </span>
                        )}
                      </p>
                      <p className="truncate text-xs text-muted-foreground">{s.email}</p>
                    </div>
                    <div className="flex shrink-0 gap-1">
                      <button
                        disabled={!canManage}
                        title={canManage ? "Reset password" : "Admin only"}
                        onClick={() => void handleReset(s)}
                        className={`rounded-lg border border-input px-2 py-1 text-xs ${
                          canManage ? "text-foreground hover:bg-muted" : "cursor-not-allowed text-muted-foreground/50"
                        }`}
                      >
                        Reset
                      </button>
                      <button
                        disabled={!canManage}
                        title={canManage ? "Delete account" : "Admin only"}
                        onClick={() => void handleDelete(s)}
                        className={`rounded-lg border border-input px-2 py-1 text-xs ${
                          canManage
                            ? "text-destructive hover:bg-destructive/10"
                            : "cursor-not-allowed text-muted-foreground/50"
                        }`}
                      >
                        Delete
                      </button>
                    </div>
                  </li>
                ))}
              </ul>
            )}

            {canManage && (
              <div className="mt-4 border-t border-border pt-4">
                {showCreate ? (
                  <CreateStaffForm
                    account={account}
                    sessionPassword={sessionPassword}
                    onDone={() => {
                      setShowCreate(false);
                      void refresh();
                    }}
                    onCancel={() => setShowCreate(false)}
                  />
                ) : (
                  <button
                    onClick={() => setShowCreate(true)}
                    className="w-full rounded-lg border border-dashed border-input py-2 text-sm font-medium text-primary hover:bg-primary/10"
                  >
                    + Create user
                  </button>
                )}
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}

function CreateStaffForm({
  account,
  sessionPassword,
  onDone,
  onCancel,
}: {
  account: StaffAccount;
  sessionPassword: string;
  onDone: () => void;
  onCancel: () => void;
}) {
  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [role, setRole] = useState<"staff" | "admin">("staff");
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!supabase) return;
    setBusy(true);
    setError(null);
    const { error: rpcError } = await supabase.rpc("staff_create_user", {
      p_admin_email: account.email,
      p_admin_password: sessionPassword,
      p_new_email: email.trim().toLowerCase(),
      p_new_password: password,
      p_new_name: name.trim(),
      p_new_role: role,
    });
    setBusy(false);
    if (rpcError) {
      setError(rpcError.message.toLowerCase().includes("duplicate") ? "That email is already in use." : rpcError.message);
      return;
    }
    onDone();
  };

  return (
    <form onSubmit={handleSubmit} className="space-y-2">
      <input
        value={name}
        onChange={(e) => setName(e.target.value)}
        placeholder="Name"
        required
        className="w-full rounded-lg border border-input bg-background px-3 py-2 text-sm text-foreground outline-none focus:ring-2 focus:ring-ring"
      />
      <input
        type="email"
        value={email}
        onChange={(e) => setEmail(e.target.value)}
        placeholder="Email"
        required
        className="w-full rounded-lg border border-input bg-background px-3 py-2 text-sm text-foreground outline-none focus:ring-2 focus:ring-ring"
      />
      <input
        type="password"
        value={password}
        onChange={(e) => setPassword(e.target.value)}
        placeholder="Password"
        required
        minLength={6}
        className="w-full rounded-lg border border-input bg-background px-3 py-2 text-sm text-foreground outline-none focus:ring-2 focus:ring-ring"
      />
      <select
        value={role}
        onChange={(e) => setRole(e.target.value as "staff" | "admin")}
        className="w-full rounded-lg border border-input bg-background px-3 py-2 text-sm text-foreground outline-none focus:ring-2 focus:ring-ring"
      >
        <option value="staff">Staff</option>
        <option value="admin">Admin</option>
      </select>
      {error && <p className="text-xs text-destructive">{error}</p>}
      <div className="flex gap-2">
        <button
          type="submit"
          disabled={busy}
          className="flex-1 rounded-lg bg-primary py-2 text-sm font-medium text-primary-foreground disabled:opacity-60"
        >
          {busy ? "Creating…" : "Create"}
        </button>
        <button type="button" onClick={onCancel} className="rounded-lg border border-input px-3 py-2 text-sm text-foreground">
          Cancel
        </button>
      </div>
    </form>
  );
}
