const fs = require("fs");
const { createClient } = require("@supabase/supabase-js");

const env = fs.readFileSync("d:/website/icount/.env", "utf8");
const url = env.match(/VITE_SUPABASE_URL=(.*)/)[1].trim();
const anonKey = env.match(/VITE_SUPABASE_ANON_KEY=(.*)/)[1].trim();

const items = require("d:/website/icount/src/data/inventory.json");

async function main() {
  const supabase = createClient(url, anonKey);
  const now = new Date().toISOString();
  const rows = items.map((it) => ({
    id: it.id,
    name: it.name,
    loc: it.loc,
    qty: it.qty,
    cost: it.cost,
    price: it.price,
    source: "base",
    updated_at: now,
  }));

  const BATCH = 1000;
  let done = 0;
  for (let i = 0; i < rows.length; i += BATCH) {
    const batch = rows.slice(i, i + BATCH);
    const { error } = await supabase.from("stock_items").upsert(batch);
    if (error) throw new Error(`batch starting at ${i} failed: ${error.message}`);
    done += batch.length;
    console.log(`upserted ${done}/${rows.length}`);
  }

  const { count, error: countErr } = await supabase
    .from("stock_items")
    .select("*", { count: "exact", head: true });
  if (countErr) throw countErr;
  console.log("final stock_items count:", count);
}

main().catch((e) => {
  console.error("FAILED:", e.message || e);
  process.exit(1);
});
