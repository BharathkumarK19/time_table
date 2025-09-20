let faculties = [];

function expandAll(open) {
    document.querySelectorAll('.faculty-collapse').forEach(el => {
        const c = new bootstrap.Collapse(el, { toggle: false });
        open ? c.show() : c.hide();
    });
}

function addFaculty() {
    faculties.push({
        Name: "",
        Full_Name: "",
        Designation: "Professor",
        Shift: "8-3",
        Weekly_Hours: 18,
        Subjects: []
    });
    render();
}

function removeFaculty(idx) {
    faculties.splice(idx, 1);
    render();
}

function addSubject(idx, type) {
    const subjBase = {
        Type: type,
        Semester: "3",
        Division: "A",
        Div_Shift: "8-3",
        Subject: "",
        Course_Code: "",
        Holidays: [],
        Num_Holidays: 0
    };
    if (type === "Theory") {
        subjBase.Theory_Classes = 3;
    } else {
        subjBase.Num_Labs = 1;
        subjBase.Batches = ["B1"];
        subjBase.Batches_Grouped = false;
    }
    faculties[idx].Subjects.push(subjBase);
    render();
}

function removeSubject(fidx, sidx) {
    faculties[fidx].Subjects.splice(sidx, 1);
    render();
}

function updateFacultyField(idx, field, value) {
    if (field === "Weekly_Hours") {
        faculties[idx][field] = parseInt(value || "0", 10);
    } else {
        faculties[idx][field] = value;
    }
}

function updateSubjectField(fidx, sidx, field, value) {
    const subj = faculties[fidx].Subjects[sidx];
    if (field === "Theory_Classes" || field === "Num_Labs" || field === "Num_Holidays") {
        subj[field] = parseInt(value || "0", 10);
    } else if (field === "Batches") {
        const parts = (value || "").split(/[, ]+/).map(x => x.trim()).filter(Boolean);
        subj.Batches = parts.length ? parts : ["B1"];
    } else if (field === "Batches_Grouped") {
        subj[field] = !!value;
    } else {
        subj[field] = value;
    }
}

function toggleHoliday(fidx, sidx, day, checked, event) {
    let subj = faculties[fidx].Subjects[sidx];
    subj.Holidays = subj.Holidays || [];
    subj.Num_Holidays = parseInt(subj.Num_Holidays || 0);

    if (checked) {
        if (subj.Holidays.length < subj.Num_Holidays) {
            if (!subj.Holidays.includes(day)) {
                subj.Holidays.push(day);
            }
        } else {
            alert("You can only select " + subj.Num_Holidays + " holidays.");
            event.target.checked = false;
        }
    } else {
        subj.Holidays = subj.Holidays.filter(d => d !== day);
    }
}

function subjectCard(fidx, sidx, s) {
    const isTheory = s.Type === "Theory";
    return `
    <div class="card subject-card mb-2">
      <div class="card-body">
        <div class="d-flex justify-content-between">
          <div class="fw-semibold">${isTheory ? "Theory" : "Lab"} #${sidx+1}</div>
          <button type="button" class="btn btn-sm btn-outline-danger" onclick="removeSubject(${fidx}, ${sidx})">Remove</button>
        </div>
        <div class="row g-2 mt-2">
          <div class="col-sm-2">
            <label class="small-label">Semester</label>
            <select class="form-select" onchange="updateSubjectField(${fidx},${sidx},'Semester', this.value)">
              <option ${s.Semester==='3'?'selected':''}>3</option>
              <option ${s.Semester==='5'?'selected':''}>5</option>
              <option ${s.Semester==='7'?'selected':''}>7</option>
            </select>
          </div>
          <div class="col-sm-2">
            <label class="small-label">Division</label>
            <input class="form-control" value="${s.Division||''}" oninput="updateSubjectField(${fidx},${sidx},'Division', this.value)">
          </div>
          <div class="col-sm-2">
            <label class="small-label">Div Shift</label>
            <select class="form-select" onchange="updateSubjectField(${fidx},${sidx},'Div_Shift', this.value)">
              <option value="8-3" ${s.Div_Shift==='8-3'?'selected':''}>8-3</option>
              <option value="10-5" ${s.Div_Shift==='10-5'?'selected':''}>10-5</option>
            </select>
          </div>
          <div class="col-sm-3">
            <label class="small-label">Subject</label>
            <input class="form-control" value="${s.Subject||''}" oninput="updateSubjectField(${fidx},${sidx},'Subject', this.value)">
          </div>
          <div class="col-sm-3">
            <label class="small-label">Course Code</label>
            <input class="form-control" value="${s.Course_Code||''}" oninput="updateSubjectField(${fidx},${sidx},'Course_Code', this.value)">
          </div>
          ${isTheory ? `
          <div class="col-sm-3">
            <label class="small-label">Theory Classes / week</label>
            <input type="number" min="1" class="form-control" value="${s.Theory_Classes||1}" oninput="updateSubjectField(${fidx},${sidx},'Theory_Classes', this.value)">
          </div>
          ` : `
          <div class="col-sm-3">
            <label class="small-label"># Labs / week</label>
            <input type="number" min="1" class="form-control" value="${s.Num_Labs||1}" oninput="updateSubjectField(${fidx},${sidx},'Num_Labs', this.value)">
          </div>
          <div class="col-sm-6">
            <label class="small-label">Batches (comma/space separated) — or tick “Group with ‘/’”</label>
            <input class="form-control" value="${(s.Batches||[]).join(', ')}" oninput="updateSubjectField(${fidx},${sidx},'Batches', this.value)">
          </div>
          <div class="col-sm-3 d-flex align-items-end">
            <div class="form-check">
              <input class="form-check-input" type="checkbox" ${s.Batches_Grouped?'checked':''} onchange="updateSubjectField(${fidx},${sidx},'Batches_Grouped', this.checked)">
              <label class="form-check-label">Group with “/”</label>
            </div>
          </div>
          `}
          <div class="col-sm-3">
            <label class="small-label"># Holidays</label>
            <input type="number" min="0" max="6" class="form-control" 
              value="${s.Num_Holidays || 0}" 
              oninput="updateSubjectField(${fidx},${sidx},'Num_Holidays', this.value)">
          </div>
          <div class="col-sm-9">
            <label class="small-label">Select Holidays</label><br>
            ${["Mon","Tue","Wed","Thu","Fri","Sat"].map(day => `
              <div class="form-check form-check-inline">
                <input class="form-check-input" type="checkbox"
                  ${s.Holidays && s.Holidays.includes(day) ? "checked" : ""}
                  onchange="toggleHoliday(${fidx}, ${sidx}, '${day}', this.checked, event)">
                <label class="form-check-label">${day}</label>
              </div>
            `).join("")}
          </div>
        </div>
      </div>
    </div>
    `;
}

function facultyCard(f, idx) {
    const collapseId = `fac-${idx}`;
    return `
    <div class="card mb-3">
      <div class="card-header d-flex justify-content-between align-items-center">
        <div class="fw-semibold">Faculty #${idx + 1} — ${f.Name || '(unnamed)'}</div>
        <div class="d-flex gap-2">
          <button type="button" class="btn btn-sm btn-outline-success" onclick="addSubject(${idx}, 'Theory')">+ Theory</button>
          <button type="button" class="btn btn-sm btn-outline-success" onclick="addSubject(${idx}, 'Lab')">+ Lab</button>
          <button type="button" class="btn btn-sm btn-outline-danger" onclick="removeFaculty(${idx})">Remove</button>
          <button class="btn btn-sm btn-outline-secondary" type="button" data-bs-toggle="collapse" data-bs-target="#${collapseId}">Toggle</button>
        </div>
      </div>
      <div id="${collapseId}" class="faculty-collapse collapse show">
        <div class="card-body">
          <div class="row g-2">
            <div class="col-md-2">
              <label class="small-label">Short Name</label>
              <input class="form-control" value="${f.Name||''}" oninput="updateFacultyField(${idx},'Name',this.value)">
            </div>
            <div class="col-md-3">
              <label class="small-label">Full Name</label>
              <input class="form-control" value="${f.Full_Name||''}" oninput="updateFacultyField(${idx},'Full_Name',this.value)">
            </div>
            <div class="col-md-3">
              <label class="small-label">Designation</label>
              <select class="form-select" onchange="updateFacultyField(${idx},'Designation',this.value)">
                <option ${f.Designation==='Professor'?'selected':''}>Professor</option>
                <option ${f.Designation==='Assistant Professor'?'selected':''}>Assistant Professor</option>
                <option ${f.Designation==='Jr Assistant Professor'?'selected':''}>Jr Assistant Professor</option>
              </select>
            </div>
            <div class="col-md-2">
              <label class="small-label">Shift</label>
              <select class="form-select" onchange="updateFacultyField(${idx},'Shift',this.value)">
                <option value="8-3" ${f.Shift==='8-3'?'selected':''}>8-3</option>
                <option value="10-5" ${f.Shift==='10-5'?'selected':''}>10-5</option>
              </select>
            </div>
            <div class="col-md-2">
              <label class="small-label">Weekly Hours</label>
              <input type="number" min="1" class="form-control" value="${f.Weekly_Hours||18}" oninput="updateFacultyField(${idx},'Weekly_Hours',this.value)">
            </div>
          </div>
          <div class="mt-3">
            ${f.Subjects.map((s, sidx) => subjectCard(idx, sidx, s)).join("")}
            ${f.Subjects.length === 0 ? '<div class="text-muted">No subjects added yet. Use the buttons above.</div>' : ''}
          </div>
        </div>
      </div>
    </div>`;
}

function render() {
    const root = document.getElementById('facultiesContainer');
    root.innerHTML = faculties.map((f, idx) => facultyCard(f, idx)).join("");
}

// ... (all other functions are the same)

async function submitAll() {
    const errorBox = document.getElementById('errorBox');
    const btn = document.getElementById('submitBtn');
    const spinner = document.getElementById('spinner');

    errorBox.classList.add('d-none');
    btn.disabled = true;
    spinner.classList.remove('d-none');

    try {
        if (faculties.length === 0) {
            throw new Error("Add at least one faculty.");
        }
        for (const f of faculties) {
            if (!f.Name) throw new Error("Every faculty needs a Short Name.");
            if (!['8-3', '10-5'].includes(f.Shift)) throw new Error(`Faculty ${f.Name} has invalid Shift.`);
            for (const s of f.Subjects) {
                if (s.Holidays.length > s.Num_Holidays) {
                    throw new Error(`Subject "${s.Subject}" for faculty ${f.Name} has more selected holidays than allowed.`);
                }
            }
        }
        
        // --- START: NEW CODE BLOCK ---
        const university = document.getElementById('university').value;
        const department = document.getElementById('department').value;
        const academic = document.getElementById('academic').value; // Make sure you have this ID in your HTML
        // --- END: NEW CODE BLOCK ---

        const res = await fetch("/generate", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            // --- START: UPDATED body payload ---
            body: JSON.stringify({ 
                university,
                department,
                academic,
                faculties
            })
            // --- END: UPDATED body payload ---
        });
        const data = await res.json();
        if (!data.ok) {
            throw new Error(data.error || "Generation failed.");
        }
        window.location = data.redirect;
    } catch (err) {
        errorBox.textContent = err.message || String(err);
        errorBox.classList.remove('d-none');
    } finally {
        spinner.classList.add('d-none');
        btn.disabled = false;
    }
}
render();
