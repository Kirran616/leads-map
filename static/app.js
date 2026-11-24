const PAGE_SIZE = 25;
let allLeads = [];
let page = 0;

const $ = (id) => document.getElementById(id);
const form = $("search-form");

function show(el, on) { el.hidden = !on; }

function showError(msg) {
  $("error").textContent = msg;
  show($("error"), Boolean(msg));
}

function visibleLeads() {
  return $("only-contact").checked
    ? allLeads.filter((l) => l.phone || l.website)
    : allLeads;
}

function cell(text) {
  const td = document.createElement("td");
  if (text) td.textContent = text;
  else { td.textContent = "–"; td.className = "dash"; }
  return td;
}

function safeUrl(raw) {
  try {
    const u = new URL(/^https?:\/\//i.test(raw) ? raw : "https://" + raw);
    return ["http:", "https:"].includes(u.protocol) ? u.href : null;
  } catch { return null; }
}

function render() {
  const leads = visibleLeads();
  const pages = Math.max(1, Math.ceil(leads.length / PAGE_SIZE));
  page = Math.min(page, pages - 1);
  const tbody = $("rows");
  tbody.replaceChildren();

  leads.slice(page * PAGE_SIZE, (page + 1) * PAGE_SIZE).forEach((l) => {
    const tr = document.createElement("tr");

    const name = document.createElement("td");
    const strong = document.createElement("strong");
    strong.textContent = l.name;
    name.append(strong);
    if (l.category) {
      const c = document.createElement("span");
      c.className = "cat";
      c.textContent = l.category;
      name.append(c);
    }
    tr.append(name, cell(l.address), cell(l.phone));

    const web = document.createElement("td");
    const href = l.website && safeUrl(l.website);
    if (href) {
      const a = document.createElement("a");
      a.href = href; a.target = "_blank"; a.rel = "noopener noreferrer";
      a.textContent = l.website.replace(/^https?:\/\//i, "").replace(/\/$/, "");
      web.append(a);
    } else { web.textContent = "–"; web.className = "dash"; }
    tr.append(web);

    const idCell = document.createElement("td");
    const idLink = document.createElement("a");
    idLink.href = "https://www.openstreetmap.org/" + l.osm_id;
    idLink.target = "_blank"; idLink.rel = "noopener noreferrer";
    idLink.textContent = l.osm_id;
    idCell.append(idLink);
    tr.append(idCell);

    tbody.append(tr);
  });

  show($("table-wrap"), leads.length > 0);
  show($("pager"), pages > 1);
  $("page-label").textContent = `Page ${page + 1} of ${pages}`;
  $("prev").disabled = page === 0;
  $("next").disabled = page >= pages - 1;
  $("export-csv").disabled = $("export-xlsx").disabled = leads.length === 0;
}

form.addEventListener("submit", async (e) => {
  e.preventDefault();
  showError("");
  const keyword = $("keyword").value.trim();
  const location = $("location").value.trim();
  if (!keyword || !location) {
    showError("Enter both a business type and a city.");
    return;
  }

  show($("loading"), true);
  show($("empty"), false);
  show($("table-wrap"), false);
  show($("pager"), false);
  $("search-btn").disabled = true;

  try {
    const res = await fetch("/api/search", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ keyword, location, limit: Number($("limit").value) }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Something went wrong.");

    allLeads = data.leads;
    page = 0;
    $("result-title").textContent = `${data.total} business${data.total === 1 ? "" : "es"} found`;
    $("result-meta").textContent =
      `${keyword} in ${data.place}` +
      (data.duplicates_removed ? ` – ${data.duplicates_removed} duplicate${data.duplicates_removed === 1 ? "" : "s"} removed` : "");
    if (data.total === 0) {
      show($("empty"), true);
      $("empty").querySelector("h2").textContent = "No businesses found";
      $("empty").querySelector("p").textContent = "Try a broader keyword such as “restaurant”, or a larger city.";
    }
    render();
  } catch (err) {
    showError(err.message);
    show($("empty"), allLeads.length === 0);
  } finally {
    show($("loading"), false);
    $("search-btn").disabled = false;
  }
});

async function exportAs(kind) {
  showError("");
  try {
    const res = await fetch(`/api/export/${kind}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ leads: visibleLeads() }),
    });
    if (!res.ok) throw new Error((await res.json()).error || "Export failed.");
    const blob = await res.blob();
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `leads.${kind}`;
    a.click();
    URL.revokeObjectURL(a.href);
  } catch (err) { showError(err.message); }
}

$("export-csv").addEventListener("click", () => exportAs("csv"));
$("export-xlsx").addEventListener("click", () => exportAs("xlsx"));
$("only-contact").addEventListener("change", () => { page = 0; render(); });
$("prev").addEventListener("click", () => { page--; render(); });
$("next").addEventListener("click", () => { page++; render(); });
