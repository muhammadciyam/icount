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

const items = inventory as Item[];
const STORAGE_KEY = "stock-count-v1";

type CountState = Record<string, number>;

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
        content: "Search items, tick counted stock and see variance instantly.",
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
  const [filter, setFilter] = useState<"all" | "pending" | "counted">("all");
  const [loaded, setLoaded] = useState(false);

  useEffect(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) setCounts(JSON.parse(raw));
    } catch {
      /* ignore */
    }
    setLoaded(true);
  }, []);

  useEffect(() => {
    if (loaded) localStorage.setItem(STORAGE_KEY, JSON.stringify(counts));
  }, [counts, loaded]);

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
            <label className="flex items-center gap-2 text-muted-foreground">
              <input
                type="checkbox"
                checked={onlyPending}
                onChange={(e) => setOnlyPending(e.target.checked)}
                className="size-4 accent-current"
              />
              Show only not-counted
            </label>
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
