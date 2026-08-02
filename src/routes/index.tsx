import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import inventory from "@/data/inventory.json";
import { supabase, supabaseEnabled } from "@/lib/supabase";

type Item = {
  id: string;
  loc: string;
  name: string;
  qty: number;
  cost: number;
  price: number;
  // "base" = part of the original bundled catalog, "custom" = added by staff.
  // Only present when items are loaded from Supabase — see supabase/schema.sql.
  source?: "base" | "custom";
};

const STORAGE_KEY = "stock-count-v1";
const INVENTORY_KEY = "stock-inventory-v1";
const AUTH_KEY = "stock-auth-v1";

// Local fallback admin login, used only when Supabase isn't configured (no
// database to store real accounts in). Once Supabase is set up, accounts
// live in the `staff_accounts` table — see supabase/schema.sql — and staff
// can be created/deleted/reset from the "Staff" panel in the app.
const FALLBACK_ADMIN = { email: "siyante003@gmail.com", password: "229022#" };

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

async function parseSheet(file: File): Promise<Item[]> {
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
    });
  });
  return out;
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
  const [query, setQuery] = useState("");
  const [counts, setCounts] = useState<CountState>({});
  const [customItems, setCustomItems] = useState<InventoryState>({});
  const [filter, setFilter] = useState<"all" | "pending" | "counted">("all");
  const [showAdd, setShowAdd] = useState(false);
  const [showStaffPanel, setShowStaffPanel] = useState(false);
  const [loaded, setLoaded] = useState(false);
  const [account, setAccount] = useState<StaffAccount | null>(null);
  // In-memory only (never persisted) — lets admin actions in the staff
  // panel skip re-prompting for a password within the same page session.
  const [sessionPassword, setSessionPassword] = useState<string | null>(null);
  const [authChecked, setAuthChecked] = useState(false);

  // When Supabase is configured, the whole catalog (base + added) lives in
  // the `stock_items` table and is loaded into `customItems` below — the
  // bundled inventory.json is only used as an offline/no-database fallback.
  const items = useMemo<Item[]>(() => {
    if (supabaseEnabled) return Object.values(customItems);
    return [
      ...(inventory as Item[]).map((it) => ({ ...it, source: "base" as const })),
      ...Object.values(customItems),
    ];
  }, [customItems]);

  // Initial load: from Supabase (shared, cross-device) when configured,
  // otherwise from this browser's local storage only.
  useEffect(() => {
    async function load() {
      if (supabaseEnabled && supabase) {
        try {
          const { data, error } = await supabase.from("stock_counts").select("item_id, qty");
          if (error) throw error;
          setCounts(Object.fromEntries((data ?? []).map((r) => [r.item_id as string, Number(r.qty)])));
        } catch (err) {
          console.error("Failed to load shared counts", err);
        }
        try {
          const { data, error } = await supabase.from("stock_items").select("*");
          if (error) throw error;
          setCustomItems(
            Object.fromEntries(
              (data ?? []).map((r) => [
                r.id as string,
                {
                  id: r.id as string,
                  name: r.name as string,
                  loc: (r.loc as string) ?? "",
                  qty: Number(r.qty),
                  cost: Number(r.cost),
                  price: Number(r.price),
                  source: (r.source as "base" | "custom" | undefined) ?? "custom",
                } satisfies Item,
              ]),
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

    const channel = supabase
      .channel("stock-sync")
      .on(
        "postgres_changes",
        { event: "*", schema: "public", table: "stock_counts" },
        (payload) => {
          setCounts((c) => {
            const next = { ...c };
            if (payload.eventType === "DELETE") {
              delete next[payload.old.item_id as string];
            } else {
              next[payload.new.item_id as string] = Number(payload.new.qty);
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
              delete next[payload.old.id as string];
            } else {
              const r = payload.new;
              next[r.id as string] = {
                id: r.id as string,
                name: r.name as string,
                loc: (r.loc as string) ?? "",
                qty: Number(r.qty),
                cost: Number(r.cost),
                price: Number(r.price),
                source: (r.source as "base" | "custom" | undefined) ?? "custom",
              };
            }
            return next;
          });
        },
      )
      .subscribe();

    return () => {
      supabase.removeChannel(channel);
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

  const saveCount = async (itemId: string, qty: number) => {
    setCounts((c) => ({ ...c, [itemId]: qty }));
    if (supabaseEnabled && supabase) {
      const { error } = await supabase
        .from("stock_counts")
        .upsert({ item_id: itemId, qty, counted_by: account?.email ?? null, updated_at: new Date().toISOString() });
      if (error) {
        console.error("Failed to save count", error);
        alert("Could not save your count to the shared database — check your connection and try again.");
      }
    }
  };

  const clearCount = async (itemId: string) => {
    setCounts((c) => {
      const n = { ...c };
      delete n[itemId];
      return n;
    });
    if (supabaseEnabled && supabase) {
      const { error } = await supabase.from("stock_counts").delete().eq("item_id", itemId);
      if (error) {
        console.error("Failed to clear count", error);
        alert("Could not update the shared database — check your connection and try again.");
      }
    }
  };

  const resetAllCounts = async () => {
    setCounts({});
    if (supabaseEnabled && supabase) {
      const { error } = await supabase.from("stock_counts").delete().neq("item_id", "");
      if (error) {
        console.error("Failed to reset counts", error);
        alert("Could not reset the shared database — check your connection and try again.");
      }
    }
  };

  const addItems = async (newItems: Item[]) => {
    setCustomItems((prev) => {
      const next = { ...prev };
      for (const it of newItems) next[it.id] = it;
      return next;
    });
    if (supabaseEnabled && supabase) {
      const { error } = await supabase.from("stock_items").upsert(
        newItems.map((it) => ({
          id: it.id,
          name: it.name,
          loc: it.loc,
          qty: it.qty,
          cost: it.cost,
          price: it.price,
          updated_at: new Date().toISOString(),
        })),
      );
      if (error) {
        console.error("Failed to save items", error);
        alert("Could not save item(s) to the shared database — check your connection and try again.");
      }
    }
  };

  const deleteItem = async (itemId: string) => {
    if (customItems[itemId]?.source === "base") return; // defense in depth — UI already hides this case
    setCustomItems((prev) => {
      const next = { ...prev };
      delete next[itemId];
      return next;
    });
    setCounts((c) => {
      const n = { ...c };
      delete n[itemId];
      return n;
    });
    if (supabaseEnabled && supabase) {
      const { error } = await supabase.from("stock_items").delete().eq("id", itemId).eq("source", "custom");
      if (error) {
        console.error("Failed to delete item", error);
        alert("Could not delete that item from the shared database — check your connection and try again.");
      }
      const { error: countError } = await supabase.from("stock_counts").delete().eq("item_id", itemId);
      if (countError) console.error("Failed to delete count for item", countError);
    }
  };

  // Clears everything added via "+ Add item" or Excel import — i.e. rows
  // tagged source='custom'. Rows tagged source='base' (the original bundled
  // catalog, now also stored in stock_items) are never touched by this.
  const deleteAllCustomItems = async () => {
    const ids = Object.entries(customItems)
      .filter(([, it]) => it.source !== "base")
      .map(([id]) => id);
    setCustomItems((prev) => {
      const next = { ...prev };
      for (const id of ids) delete next[id];
      return next;
    });
    setCounts((c) => {
      const n = { ...c };
      for (const id of ids) delete n[id];
      return n;
    });
    if (supabaseEnabled && supabase) {
      const { error } = await supabase.from("stock_items").delete().eq("source", "custom");
      if (error) {
        console.error("Failed to delete all items", error);
        alert("Could not delete all items from the shared database — check your connection and try again.");
      }
      if (ids.length > 0) {
        const { error: countError } = await supabase.from("stock_counts").delete().in("item_id", ids);
        if (countError) console.error("Failed to delete counts for items", countError);
      }
    }
  };

  const results = useMemo(() => {
    const q = query.trim().toLowerCase();
    let list = items;
    if (q) {
      const terms = q.split(/\s+/);
      list = list.filter((i) => {
        const hay = `${i.id} ${i.name}`.toLowerCase();
        return terms.every((t) => hay.includes(t));
      });
    }
    if (filter === "pending") list = list.filter((i) => counts[i.id] === undefined);
    if (filter === "counted") list = list.filter((i) => counts[i.id] !== undefined);
    return list.slice(0, 200);
  }, [query, filter, counts]);

  const done = Object.keys(counts).length;
  const pct = Math.round((done / items.length) * 100);

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

  return (
    <main className="min-h-screen bg-background pb-24">
      <header className="sticky top-0 z-10 border-b border-border bg-card/95 shadow-sm backdrop-blur">
        <div className="mx-auto max-w-3xl px-4 py-4 sm:px-6">
          <div className="flex items-center justify-between gap-3">
            <div className="flex items-center gap-2.5">
              <div className="flex size-9 items-center justify-center rounded-xl bg-primary text-base font-bold text-primary-foreground shadow-sm">
                7M
              </div>
              <div>
                <h1 className="text-base font-semibold leading-tight tracking-tight text-foreground sm:text-lg">
                  Stock Count
                </h1>
                <p className="text-xs text-muted-foreground">
                  {done} / {items.length} counted · {pct}%
                </p>
              </div>
            </div>
            <div className="flex items-center gap-2">
              <span className="hidden text-xs text-muted-foreground sm:inline">
                Hi, <span className="font-medium text-foreground">{account.name}</span>
                {account.role === "admin" && (
                  <span className="ml-1.5 rounded-full bg-primary/10 px-1.5 py-0.5 text-[10px] font-semibold text-primary">
                    Admin
                  </span>
                )}
              </span>
              <button
                onClick={() => setShowStaffPanel(true)}
                className="rounded-lg border border-input px-3 py-2 text-xs font-medium text-muted-foreground transition-colors hover:bg-muted hover:text-foreground"
              >
                Staff
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
                className="rounded-lg border border-input px-3 py-2 text-xs font-medium text-muted-foreground transition-colors hover:bg-muted hover:text-foreground"
              >
                Log out
              </button>
            </div>
          </div>
          <div className="mt-3 h-2 w-full overflow-hidden rounded-full bg-muted">
            <div
              className="h-full rounded-full bg-primary transition-all duration-300"
              style={{ width: `${pct}%` }}
            />
          </div>
          <input
            autoFocus
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            inputMode="search"
            placeholder="Search item name or barcode / ID…"
            className="mt-3 w-full rounded-xl border border-input bg-background px-4 py-3.5 text-base text-foreground outline-none placeholder:text-muted-foreground focus:ring-2 focus:ring-ring sm:py-3"
          />
          <div className="mt-3 flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
            <div className="flex gap-1 rounded-lg border border-input p-1">
              {(["all", "pending", "counted"] as const).map((f) => (
                <button
                  key={f}
                  onClick={() => setFilter(f)}
                  className={`flex-1 rounded-md px-3 py-2 text-xs font-medium transition-colors sm:flex-none sm:px-2.5 sm:py-1 ${
                    filter === f
                      ? "bg-primary text-primary-foreground"
                      : "text-muted-foreground hover:text-foreground"
                  }`}
                >
                  {f === "all" ? "All" : f === "pending" ? "Not counted" : "Counted"}
                </button>
              ))}
            </div>
            <div className="flex flex-wrap items-center gap-2 text-sm">
              <label className="flex min-h-9 cursor-pointer items-center rounded-lg px-2.5 py-1.5 font-medium text-primary transition-colors hover:bg-primary/10">
                Import Excel
                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  className="hidden"
                  onChange={async (e) => {
                    const file = e.target.files?.[0];
                    e.target.value = "";
                    if (!file) return;
                    try {
                      const imported = await parseSheet(file);
                      if (imported.length === 0) {
                        alert("No rows found. Make sure the sheet has columns like Barcode/ID, Name, Qty.");
                        return;
                      }
                      await addItems(imported);
                      alert(`Imported ${imported.length} items.`);
                    } catch (err) {
                      alert("Could not read that file. Please upload a valid Excel or CSV file.");
                      console.error(err);
                    }
                  }}
                />
              </label>
              <button
                onClick={() => setShowAdd(true)}
                className="min-h-9 rounded-lg px-2.5 py-1.5 font-medium text-primary transition-colors hover:bg-primary/10"
              >
                + Add item
              </button>

              <button
                onClick={() => {
                  if (account.role !== "admin") {
                    alert("Only an admin can reset the count.");
                    return;
                  }
                  if (confirm("Clear all counted items?")) void resetAllCounts();
                }}
                title={account.role === "admin" ? undefined : "Admin only"}
                className={`min-h-9 rounded-lg px-2.5 py-1.5 font-medium transition-colors ${
                  account.role === "admin"
                    ? "text-destructive hover:bg-destructive/10"
                    : "cursor-not-allowed text-muted-foreground/50"
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
                  if (Object.keys(customItems).length === 0) {
                    alert("No added/imported items to delete.");
                    return;
                  }
                  if (
                    confirm(
                      `Delete all ${Object.keys(customItems).length} added/imported item(s)? The base inventory list isn't affected.`,
                    )
                  ) {
                    void deleteAllCustomItems();
                  }
                }}
                title={account.role === "admin" ? undefined : "Admin only"}
                className={`min-h-9 rounded-lg px-2.5 py-1.5 font-medium transition-colors ${
                  account.role === "admin"
                    ? "text-destructive hover:bg-destructive/10"
                    : "cursor-not-allowed text-muted-foreground/50"
                }`}
              >
                Delete all items
              </button>
            </div>
          </div>
        </div>
      </header>

      <div className="mx-auto max-w-3xl px-3 sm:px-6">
        {results.length === 0 ? (
          <p className="py-16 text-center text-sm text-muted-foreground">No items found.</p>
        ) : (
          <ul className="mt-3 space-y-2">
            {results.map((item) => (
              <Row
                key={item.id}
                item={item}
                counted={counts[item.id]}
                isCustom={item.source !== "base"}
                isAdmin={account.role === "admin"}
                onSave={(v) => void saveCount(item.id, v)}
                onClear={() => void clearCount(item.id)}
                onDelete={() => void deleteItem(item.id)}
              />
            ))}
          </ul>
        )}
        {results.length === 200 && (
          <p className="py-4 text-center text-xs text-muted-foreground">
            Showing first 200 matches — refine your search.
          </p>
        )}
      </div>

      {showAdd && (
        <AddItemModal
          onClose={() => setShowAdd(false)}
          onSave={(item) => {
            void addItems([item]);
            setShowAdd(false);
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
    </main>
  );
}

function Row({
  item,
  counted,
  isCustom,
  isAdmin,
  onSave,
  onClear,
  onDelete,
}: {
  item: Item;
  counted?: number | undefined;
  isCustom: boolean;
  isAdmin: boolean;
  onSave: (v: number) => void;
  onClear: () => void;
  onDelete: () => void;
}) {
  const [open, setOpen] = useState(false);
  const [draft, setDraft] = useState("");
  const isDone = counted !== undefined;
  const diff = isDone ? counted - item.qty : 0;

  return (
    <li
      className={`overflow-hidden rounded-xl border transition-colors ${
        isDone ? "border-primary/30 bg-primary/5" : "border-border bg-card"
      }`}
    >
      <div className="flex items-center gap-3 px-3 py-3">
        <button
          aria-label={isDone ? "Mark as not counted" : "Mark as counted"}
          onClick={() => (isDone ? onClear() : setOpen((o) => !o))}
          className={`flex size-8 shrink-0 items-center justify-center rounded-full border text-sm transition-colors ${
            isDone
              ? "border-primary bg-primary text-primary-foreground"
              : "border-input text-muted-foreground"
          }`}
        >
          {isDone ? "✓" : ""}
        </button>
        <button onClick={() => setOpen((o) => !o)} className="min-w-0 flex-1 text-left">
          <p className="truncate text-sm font-medium text-foreground">{item.name}</p>
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
            <span>Location: <span className="text-foreground">{item.loc}</span></span>
            <span>System qty: <span className="text-foreground">{item.qty}</span></span>
            <span>Cost: <span className="text-foreground">{item.cost.toFixed(3)}</span></span>
            <span>Sale price: <span className="text-foreground">{item.price.toFixed(2)}</span></span>
          </div>
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
              className={`w-full rounded-lg border px-3 py-2 text-sm font-medium transition-colors ${
                isAdmin
                  ? "border-destructive/40 text-destructive hover:bg-destructive/10"
                  : "cursor-not-allowed border-input text-muted-foreground/50"
              }`}
            >
              Delete item
            </button>
          )}
        </div>
      )}
    </li>
  );
}

function AddItemModal({
  onClose,
  onSave,
}: {
  onClose: () => void;
  onSave: (item: Item) => void;
}) {
  const [name, setName] = useState("");
  const [loc, setLoc] = useState("");
  const [qty, setQty] = useState("");
  const [cost, setCost] = useState("");
  const [price, setPrice] = useState("");

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const n = name.trim();
    const qNum = Number(qty);
    const cNum = Number(cost);
    const pNum = Number(price);
    if (!n || Number.isNaN(qNum) || Number.isNaN(cNum) || Number.isNaN(pNum)) return;
    onSave({
      id: `custom-${Date.now()}`,
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
          <h2 className="text-lg font-semibold text-foreground">Add new item</h2>
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
            Save item
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
    // eslint-disable-next-line react-hooks/exhaustive-deps
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
