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
    setLoaded(true);
  }, []);

  useEffect(() => {
    if (loaded) localStorage.setItem(STORAGE_KEY, JSON.stringify(counts));
  }, [counts, loaded]);

  useEffect(() => {
    if (loaded) localStorage.setItem(INVENTORY_KEY, JSON.stringify(customItems));
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

  return (
    <main className="min-h-screen bg-background pb-24">
      <header className="sticky top-0 z-10 border-b border-border bg-card/95 backdrop-blur">
        <div className="mx-auto max-w-3xl px-4 py-4">
          <div className="flex items-baseline justify-between gap-3">
            <h1 className="text-lg font-semibold tracking-tight text-foreground">Stock Count</h1>
            <span className="text-sm text-muted-foreground">
              {done} / {items.length} counted ({pct}%)
            </span>
          </div>
          <div className="mt-2 h-1.5 w-full overflow-hidden rounded-full bg-muted">
            <div className="h-full bg-primary transition-all" style={{ width: `${pct}%` }} />
          </div>
          <input
            autoFocus
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            inputMode="search"
            placeholder="Search item name or barcode / ID…"
            className="mt-3 w-full rounded-xl border border-input bg-background px-4 py-3 text-base text-foreground outline-none placeholder:text-muted-foreground focus:ring-2 focus:ring-ring"
          />
          <div className="mt-2 flex items-center justify-between text-sm">
            <div className="flex gap-1 rounded-lg border border-input p-0.5">
              {(["all", "pending", "counted"] as const).map((f) => (
                <button
                  key={f}
                  onClick={() => setFilter(f)}
                  className={`rounded-md px-2.5 py-1 text-xs font-medium transition-colors ${
                    filter === f
                      ? "bg-primary text-primary-foreground"
                      : "text-muted-foreground hover:text-foreground"
                  }`}
                >
                  {f === "all" ? "All" : f === "pending" ? "Not counted" : "Counted"}
                </button>
              ))}
            </div>
            <div className="flex items-center gap-2">
              <label className="cursor-pointer rounded-md px-2 py-1 text-primary hover:bg-primary/10">
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
                className="rounded-md px-2 py-1 text-primary hover:bg-primary/10"
              >
                + Add item
              </button>

              <button
                onClick={() => {
                  if (confirm("Clear all counted items?")) setCounts({});
                }}
                className="rounded-md px-2 py-1 text-destructive hover:bg-destructive/10"
              >
                Reset count
              </button>
            </div>
          </div>
        </div>
      </header>

      <div className="mx-auto max-w-3xl px-4">
        {results.length === 0 ? (
          <p className="py-16 text-center text-sm text-muted-foreground">No items found.</p>
        ) : (
          <ul className="divide-y divide-border">
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
    <li className={isDone ? "bg-accent/40" : undefined}>
      <div className="flex items-center gap-3 px-1 py-3">
        <button
          aria-label={isDone ? "Mark as not counted" : "Mark as counted"}
          onClick={() => (isDone ? onClear() : setOpen((o) => !o))}
          className={`flex size-7 shrink-0 items-center justify-center rounded-full border text-sm ${
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
              <span className={diff === 0 ? " text-primary" : " text-destructive"}>
                {" "}
                · counted {counted} ({diff > 0 ? "+" : ""}
                {diff})
              </span>
            )}
          </p>
        </button>
      </div>

      {open && (
        <div className="space-y-3 rounded-xl bg-muted/60 px-3 py-3">
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
