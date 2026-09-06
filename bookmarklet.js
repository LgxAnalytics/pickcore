// PickCore shipment autofill bookmarklet
//
// Reads a shipment payload (JSON) from the clipboard and fills a forwarder
// portal form with it. The portal is third party, so there is no API and no
// stable markup: fields are located by their visible label, then by the text
// of the surrounding cell, then by name/id/placeholder as a last resort.
//
// Design notes:
//  - nothing is ever submitted; the operator reviews and presses Submit
//  - every filled field is marked with dataset.fwFilled so a later pass does
//    not overwrite it and so a repeat run is idempotent
//  - selects are matched on exact option text first, then on a prefix, because
//    portals localise option labels
//  - after country and parcel-count changes the form re-renders and drops
//    values, so applied fields are verified and re-applied once
//  - a summary box reports what was filled and what was not found
//
// This file is the source of truth. pickcore.py holds a percent-encoded copy
// generated from it by _build_step5.py - edit here, never there.

(async function() {
  let raw;
  try {
    raw = await navigator.clipboard.readText();
  } catch (e) {
    alert("Cannot read the clipboard.\nClick once on the page background and try again.");
    return;
  }
  let d;
  try {
    d = JSON.parse(raw);
  } catch (e) {
    alert(
      "The clipboard holds no shipment payload (JSON).\nCopy the payload from the Forwarder view first."
      );
    return;
  }
  const MAP_MAIN = [
    ["your_reference", ["your reference"]],
    ["delivery_reference", ["delivery reference"]],
    ["company_name", ["company name"]],
    ["contact_name", ["contact name"]],
    ["address1", ["address line 1", "street name"]],
    ["address2", ["address line 2"]],
    ["address3", ["address line 3"]],
    ["postcode", ["postal code", "postcode", "9999 aa"]],
    ["city", ["town", "city"]],
    ["telephone", ["telephone", "phone"]],
    ["email", ["email"]]
  ];
  const MAP_PARCEL = [
    ["description", ["parcel description"]],
    ["weight_kg", ["parcel weight"]],
    ["length_cm", ["parcel length"]],
    ["width_cm", ["parcel width"]],
    ["height_cm", ["parcel height"]],
    ["value", ["value"]]
  ];
  const norm = (s) => (s || "").toLowerCase().replace(/\s+/g, " ").trim();

  function scan() {
    return [...document.querySelectorAll("input, select, textarea")].filter((e) => e.type !==
      "hidden" && !e.disabled && e.offsetParent !== null).map((e) => {
      let lab = "";
      if (e.labels && e.labels[0]) lab = e.labels[0].innerText;
      if (!lab) {
        const cell = e.closest("td, div, tr");
        if (cell) {
          const prev = cell.previousElementSibling;
          if (prev) lab = prev.innerText;
        }
      }
      return {
        el: e,
        txt: norm([lab, e.name, e.id, e.placeholder].join(" "))
      };
    });
  }
  const findIn = (list, frags) => {
    for (const f of frags) {
      const hit = list.find((x) => x.txt.includes(f) && !x.el.dataset.fwFilled);
      if (hit) return hit;
    }
    return null;
  };
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const setValue = (el, val) => {
    const v = String(val == null ? "" : val);
    if (el.tagName === "SELECT") {
      if (!el.options || !el.options.length) return false;
      const opt = [...el.options].find((o) => norm(o.text) === norm(v) || norm(o.value) ===
        norm(v)) || [...el.options].find((o) => norm(o.text).startsWith(norm(v)));
      if (!opt) return false;
      el.value = opt.value;
    } else {
      el.focus();
      el.value = v;
    }
    el.dispatchEvent(new Event("input", {
      bubbles: true
    }));
    el.dispatchEvent(new Event("change", {
      bubbles: true
    }));
    el.dispatchEvent(new Event("blur", {
      bubbles: true
    }));
    el.style.outline = "2px solid #27c66d";
    el.dataset.fwFilled = "1";
    return true;
  };

  function tickCheckbox(labelFrag) {
    const box = [...document.querySelectorAll('input[type="checkbox"]')].find((e) => {
      let lab = e.labels && e.labels[0] ? e.labels[0].innerText : "";
      if (!lab) {
        const par = e.closest("td, div, label, tr");
        lab = par ? par.innerText : "";
      }
      return norm([lab, e.name, e.id].join(" ")).includes(labelFrag);
    });
    if (!box) return false;
    if (!box.checked) box.click();
    box.style.outline = "2px solid #27c66d";
    return box.checked;
  }

  function setCountry(want) {
    const w = norm(want);
    if (!w) return false;
    const sels = [...document.querySelectorAll("select")];
    for (const sel of sels) {
      if (!sel.options || !sel.options.length) continue;
      const opt = [...sel.options].find((o) => norm(o.text) === w || norm(o.value) === w);
      if (!opt) continue;
      if (norm(sel.value) === norm(opt.value)) {
        sel.style.outline = "2px solid #27c66d";
        return true;
      }
      sel.value = opt.value;
      sel.dispatchEvent(new Event("input", {
        bubbles: true
      }));
      sel.dispatchEvent(new Event("change", {
        bubbles: true
      }));
      try {
        if (window.jQuery) window.jQuery(sel).trigger("change");
      } catch (e) {}
      sel.style.outline = "2px solid #27c66d";
      return true;
    }
    return false;
  }

  function setPackaging(want) {
    const w = norm(want);
    const isPack = (t) => t.includes("packaging type") || t.includes("packing type");
    const sels = [...document.querySelectorAll("select")].filter((e) => isPack(norm([(e
        .labels && e.labels[0] ? e.labels[0].innerText : ""), e.name, e.id].join(" "))) ||
      isPack(norm((e.closest("td, div, tr") || {}).innerText || "")));
    for (const sel of sels) {
      if (!sel.options || !sel.options.length) continue;
      const opt = [...sel.options].find((o) => norm(o.text) === w || norm(o.value) === w);
      if (opt) {
        sel.value = opt.value;
        sel.dispatchEvent(new Event("input", {
          bubbles: true
        }));
        sel.dispatchEvent(new Event("change", {
          bubbles: true
        }));
        try {
          if (window.jQuery) window.jQuery(sel).trigger("change");
        } catch (e) {}
        sel.style.outline = "2px solid #27c66d";
        return true;
      }
    }
    const trigger = [...document.querySelectorAll("button, a, div, span")].find((e) => norm(e
        .textContent) === "select packaging type" || norm(e.textContent) ===
      "select packing type");
    if (trigger) {
      trigger.click();
      const item = [...document.querySelectorAll("li, option, div, span, a")].find((e) => norm(e
        .textContent) === w && e.offsetParent !== null);
      if (item) {
        item.click();
        return true;
      }
    }
    return false;
  }
  const filled = [],
    missed = [];
  const applied = [];

  function fillGroup(map, list) {
    for (const [key, frags] of map) {
      const val = d[key];
      if (val === undefined || val === "" || val === null) continue;
      const f = findIn(list, frags);
      if (f && setValue(f.el, val)) {
        filled.push(key);
        applied.push({
          el: f.el,
          key: key,
          val: String(val)
        });
      } else {
        missed.push(key);
      }
    }
  }
  if (d.country) {
    if (setCountry(d.country)) {
      filled.push("country");
      await sleep(800);
    } else missed.push("country");
  }
  fillGroup(MAP_MAIN, scan());
  if (d.parcels !== undefined && d.parcels !== "" && String(d.parcels) !== "1") {
    const pf = findIn(scan(), ["number of parcels"]);
    if (pf) {
      setValue(pf.el, d.parcels);
      filled.push("parcels");
      await sleep(700);
    } else {
      missed.push("parcels");
    }
  }
  fillGroup(MAP_PARCEL, scan());
  if (tickCheckbox("ready now")) filled.push("ready_now");
  else missed.push("ready_now");
  if (setPackaging(d.packing_type || "Box")) filled.push("packing_type");
  else missed.push("packing_type");
  await sleep(600);
  let repaired = 0;
  for (const a of applied) {
    try {
      if (a.el.isConnected && String(a.el.value || "") !== a.val && a.el.tagName !== "SELECT") {
        a.el.dataset.fwFilled = "";
        if (setValue(a.el, a.val)) repaired++;
      }
    } catch (e) {}
  }
  const box = document.createElement("div");
  box.style.cssText =
    "position:fixed;right:16px;bottom:16px;z-index:999999;background:#161922;color:#e9edf2;" +
    "font:12px/1.5 Segoe UI,Arial;padding:12px 14px;border:1px solid #2a2f3a;border-radius:8px;" +
    "max-width:340px;box-shadow:0 6px 24px rgba(0,0,0,.4)";
  box.innerHTML = '<b style="color:#4d8dff">Fill shipment</b><br>' +
    '<span style="color:#27c66d">filled: ' + filled.length + "</span>" + (missed.length ?
      '<br><span style="color:#f5b342">not found: ' + missed.join(", ") + "</span>" :
      '<br><span style="color:#8a93a0">all mapped fields filled</span>') + (repaired ?
      '<br><span style="color:#8a93a0">re-applied after form refresh: ' + repaired + "</span>" :
      "") + '<br><span style="color:#5a6675">Check the form and submit manually.</span>';
  document.body.appendChild(box);
  setTimeout(() => box.remove(), 9000);
})();
