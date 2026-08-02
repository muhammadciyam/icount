import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import inventory from "@/data/inventory.json";

type Item = {
  id: string;
  loc: string;
  name: string;
  qty: number;
  cost: number;
  price: number;
};

const STORAGE_KEY = "stock-count-v1";
const INVENTORY_KEY = "stock-inventory-v1";
const AUTH_KEY = "stock-auth-v1";

// Staff accounts. Add, remove, or edit entries here — each staff member gets
// their own name + password. There's no backend, so this list only lives in
// the deployed code (edit it and redeploy to change accounts).
const STAFF_ACCOUNTS: { name: string; password: string }[] = [
  { name: "Admin", password: "sevenmart2024" },
  { name: "Staff 1", password: "staff1pass" },
  { name: "Staff 2", password: "staff2pass" },
];

type CountState = Record<string, number>;
type InventoryState = Record<string, Item>;

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
  const [loaded, setLoaded] = useState(false);
  const [staffName, setStaffName] = useState<string | null>(null);
  const [authChecked, setAuthChecked] = useState(false);

  const items = useMemo<Item[]>(() => {
    return [...(inventory as Item[]), ...Object.values(customItems)];
  }, [customItems]);

  useEffect(() => {
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
    try {
      const rawName = localStorage.getItem(AUTH_KEY);
      if (rawName && STAFF_ACCOUNTS.some((s) => s.name === rawName)) setStaffName(rawName);
    } catch {
      /* ignore */
    }
    setLoaded(true);
    setAuthChecked(true);
  }, []);

  useEffect(() => {
    if (!loaded) return;
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(counts));
    } catch (err) {
      console.error("Failed to save counts", err);
      alert("Could not save your count — your device's storage may be full or unavailable.");
    }
  }, [counts, loaded]);

  useEffect(() => {
    if (!loaded) return;
    try {
      localStorage.setItem(INVENTORY_KEY, JSON.stringify(customItems));
    } catch (err) {
      console.error("Failed to save imported items", err);
      alert("Could not save your imported items — your device's storage may be full or unavailable.");
    }
  }, [customItems, loaded]);

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

  if (!staffName) {
    return (
      <LoginScreen
        onSuccess={(name) => {
          try {
            localStorage.setItem(AUTH_KEY, name);
          } catch {
            /* ignore */
          }
          setStaffName(name);
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
                Hi, <span className="font-medium text-foreground">{staffName}</span>
              </span>
              <button
                onClick={() => {
                  if (confirm("Log out?")) {
                    try {
                      localStorage.removeItem(AUTH_KEY);
                    } catch {
                      /* ignore */
                    }
                    setStaffName(null);
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
                      setCustomItems((prev) => {
                        const next = { ...prev };
                        for (const it of imported) next[it.id] = it;
                        return next;
                      });
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

              {staffName === "Admin" && (
                <button
                  onClick={() => {
                    if (confirm("Clear all counted items?")) setCounts({});
                  }}
                  className="min-h-9 rounded-lg px-2.5 py-1.5 font-medium text-destructive transition-colors hover:bg-destructive/10"
                >
                  Reset count
                </button>
              )}
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
                onSave={(v) => setCounts((c) => ({ ...c, [item.id]: v }))}
                onClear={() =>
                  setCounts((c) => {
                    const n = { ...c };
                    delete n[item.id];
                    return n;
                  })
                }
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
            setCustomItems((prev) => ({ ...prev, [item.id]: item }));
            setShowAdd(false);
          }}
        />
      )}
    </main>
  );
}

function Row({
  item,
  counted,
  onSave,
  onClear,
}: {
  item: Item;
  counted?: number | undefined;
  onSave: (v: number) => void;
  onClear: () => void;
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

function LoginScreen({ onSuccess }: { onSuccess: (name: string) => void }) {
  const [name, setName] = useState(STAFF_ACCOUNTS[0]?.name ?? "");
  const [password, setPassword] = useState("");
  const [error, setError] = useState(false);

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const account = STAFF_ACCOUNTS.find((s) => s.name === name);
    if (account && password === account.password) {
      onSuccess(account.name);
    } else {
      setError(true);
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
          <p className="mt-1 text-sm text-muted-foreground">Select your name and enter your password.</p>
        </div>
        <form onSubmit={handleSubmit} className="mt-6 space-y-3">
          <select
            value={name}
            onChange={(e) => {
              setName(e.target.value);
              setError(false);
            }}
            className="w-full rounded-xl border border-input bg-background px-4 py-3 text-base text-foreground outline-none focus:ring-2 focus:ring-ring"
          >
            {STAFF_ACCOUNTS.map((s) => (
              <option key={s.name} value={s.name}>
                {s.name}
              </option>
            ))}
          </select>
          <input
            autoFocus
            type="password"
            inputMode="text"
            value={password}
            onChange={(e) => {
              setPassword(e.target.value);
              setError(false);
            }}
            placeholder="Password"
            className={`w-full rounded-xl border bg-background px-4 py-3 text-base text-foreground outline-none placeholder:text-muted-foreground focus:ring-2 focus:ring-ring ${
              error ? "border-destructive" : "border-input"
            }`}
          />
          {error && <p className="text-sm text-destructive">Incorrect password. Try again.</p>}
          <button
            type="submit"
            className="w-full rounded-xl bg-primary py-3 text-sm font-semibold text-primary-foreground shadow-sm transition-opacity hover:opacity-90"
          >
            Log in
          </button>
        </form>
      </div>
    </main>
  );
}
