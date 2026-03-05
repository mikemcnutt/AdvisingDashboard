<script>
(function(){
  function qs(name){
    try{
      return new URLSearchParams(window.location.search).get(name) || "";
    }catch(e){
      return "";
    }
  }

  async function loadFromUrl(){
    const jsonUrl = qs("json");
    if(!jsonUrl) return;
    const res = await fetch(jsonUrl, { cache: "no-store" });
    if(!res.ok) throw new Error("Load failed: HTTP " + res.status);
    const snap = await res.json();
    applySnapshot(snap);
    localStorage.setItem("advising_snapshot_last", JSON.stringify(snap));
    if(typeof build === "function") await build(true);
  }

  async function saveToDashboard(){
    const saveUrl = qs("save");
    if(!saveUrl) return false;

    const snap = {
      v:1,
      meta:{ts:Date.now()},
      student:{firstName:$("firstName").value||"",lastName:$("lastName").value||"",studentId:$("studentId").value||"",kctcsEmail:$("kctcsEmail").value||"",personalEmail:$("personalEmail").value||"",phone:$("phone").value||""},
      selection:{scenario:scenarioEl.value||"",subplan:subplanEl.value||"",aa:!!aaToggle.checked,semesterCount:$("semesterCount")?$("semesterCount").value:"0"},
      data:{courses:window.__lastCourses||[],scores:window.__lastScores||[],manualCourses:window.__manualCourses||[],manualScores:window.__manualScores||[],notes:advisorNotesEl.value||"",overrides:window.__overrides||{},semesterPlans:window.__semesterPlans||[]}
    };

    const res = await fetch(saveUrl, {
      method: "POST",
      headers: {"Content-Type":"application/json"},
      body: JSON.stringify(snap)
    });

    if(!res.ok) throw new Error("Save failed: HTTP " + res.status);

    localStorage.setItem("advising_snapshot_last", JSON.stringify(snap));

    const savedEl = document.getElementById("lastSavedStamp");
    if(savedEl) savedEl.textContent = "Saved to Advising folder " + new Date().toLocaleString();

    return true;
  }

  document.addEventListener("DOMContentLoaded", function(){
    loadFromUrl().catch(e=>{
      console.error("Auto-load failed:", e);
      alert("Could not auto-load student JSON from the dashboard link.");
    });

    const saveUrl = qs("save");
    if(saveUrl){
      const btn = document.getElementById("saveSnapBtn");
      if(btn){
        btn.addEventListener("click", async function(e){
          e.preventDefault();
          e.stopPropagation();
          try{
            await saveToDashboard();
          }catch(err){
            console.error("Dashboard save error:", err);
            alert("Could not save back to the Advising folder.\n\n" + (err.message || err));
          }
        }, true);
      }
    }
  });
})();
</script>
