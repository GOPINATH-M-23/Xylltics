/* ================================================================
   AcadAnalytics — Dynamic Excel Analytics Engine  v3.0
   ================================================================
   MODULES (in order of data flow):
     UTILS          — shared helpers
     COLUMN_MAP     — fuzzy / regex / synonym auto-detection
     PARSER         — SheetJS → raw rows
     NORMALIZER     — raw rows → normalizedStudentData (THE source of truth)
     FAIL_DETECTOR  — ONE authoritative pass/fail function
     ANALYTICS      — pure reusable computations on normalizedStudentData
     CHART_ENGINE   — Chart.js builders, all driven by ANALYTICS output
     INSIGHTS_ENGINE— AI-style text from ANALYTICS only
     UI             — DOM updates, table, modal, filters
     APP            — orchestrator / state
     EXPORT         — download functions
   ================================================================
   GOLDEN RULE: Every module reads from APP.data (normalizedStudentData).
                Nobody reads raw rows. Nobody hardcodes column names.
   ================================================================ */

"use strict";

/* ================================================================
   UTILS
   ================================================================ */
const UTILS = {
  $: id => document.getElementById(id),

  norm: v => String(v ?? '').trim().toLowerCase().replace(/[^a-z0-9]/g, ''),

  /** Parse number safely — returns null for empty/NaN */
  n: v => {
    if (v === null || v === undefined || v === '') return null;
    const s = String(v).replace(/[^0-9.\-+]/g, '');
    const f = parseFloat(s);
    return isNaN(f) ? null : f;
  },

  /** Round to dp decimal places */
  r: (v, dp = 1) => typeof v === 'number' ? +v.toFixed(dp) : v,

  /** clamp value to [0, max] */
  clamp: (v, max) => Math.min(Math.max(v ?? 0, 0), max),

  /** Show / hide by id or element */
  show: el => (typeof el === 'string' ? UTILS.$(el) : el)?.classList.remove('hidden'),
  hide: el => (typeof el === 'string' ? UTILS.$(el) : el)?.classList.add('hidden'),

  debounce(fn, ms) { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; },

  /** Animated counter */
  animCount(id, target) {
    const el = UTILS.$(id);
    if (!el) return;
    const end = +target;
    if (isNaN(end)) { el.textContent = target; return; }
    let cur = 0;
    const step = end / (800 / 16);
    const t = setInterval(() => {
      cur = Math.min(cur + step, end);
      el.textContent = Math.floor(cur);
      if (cur >= end) clearInterval(t);
    }, 16);
  },

  toast(msg, ms = 3200) {
    const t = UTILS.$('toast');
    if (!t) return;
    UTILS.$('toastMsg').textContent = msg;
    t.classList.remove('hidden');
    clearTimeout(UTILS._tt);
    UTILS._tt = setTimeout(() => t.classList.add('hidden'), ms);
  },

  loader(msg) {
    const el = UTILS.$('loaderMsg');
    if (el) el.textContent = msg;
    UTILS.show('loaderOverlay');
  },

  loaderOff() { UTILS.hide('loaderOverlay'); },
};

/* ================================================================
   COLUMN_MAP — fuzzy, regex, synonym, keyword-score detection
   ================================================================ */
const COLUMN_MAP = (() => {

  /* --- Synonym dictionary: role → [alias list] --- */
  const DICT = {
    name: ['name','studentname','student_name','stdname','sname','fullname',
      'candidatename','pupilname','nameofstudent','nameof','student','learner'],
    roll: ['rollno','roll_no','rollnumber','roll','rollnum','rollid','htno',
      'hallticketno','regno','reg_no','regnumber','registrationno','regid',
      'enrollmentno','enrolno','admissionno','studentid','sid','sno','id'],
    subject: ['subject','subjectname','sub','subjectcode','subject_code',
      'coursename','coursecode','course','papercode','module','paper',
      'subcode','scode','subjectid'],
    department: ['department','dept','branch','stream','discipline',
      'faculty','division','program','programme'],
    section: ['section','sec','class','batch','group','classname'],
    internal: ['internal','internalmarks','internalmark','internalassessment',
      'ia','cia','sessional','ta','continuousassessment','ca','internalexam',
      'midterm','midmarks','mid','midsem','internaltotal','intmarks','intotal',
      'internalmax','totalinternal'],
    external: ['external','externalmarks','externalmark','semesterexam',
      'semexam','ese','endterm','endsem','endmarks','finalexam','fe',
      'theory','thmarks','semmarks','semestermarks','semmks','semestermark',
      'examinationmarks','universityexam','unimarks','endexam'],
    total: ['total','totalmarks','grandtotal','grand_total','totalscore',
      'finaltotal','overallmarks','overall','score','obtainedmarks','tmarks',
      'tot','totalout','summarks','aggregate','finalmarks','obtmarks'],
    grade: ['grade','lettergrade','gradetype','gradeletter','gradeval',
      'gradeobtained','gradeawarded'],
    gpa: ['gpa','cgpa','sgpa','gradepoint','grade_point','gp','gradepts',
      'gradepoints','creditpoints','qualitypoints','points','gpax'],
    result: ['result','status','passfail','pass_fail','outcome','remarks',
      'passorfail','finalstatus','examstatus','examresult','decision','verdict'],
    attendance: ['attendance','attend','att','attendancepercent',
      'attendancepercentage','presentdays','absentdays','attendancemarks'],
    cat1: ['cat1','cat1marks','ca1','ct1','ut1','test1','t1','exam1',
      'component1','comp1','internal1','int1','firsttest','periodicaltest1',
      'pt1','periodical1','cia1'],
    cat2: ['cat2','cat2marks','ca2','ct2','ut2','test2','t2','exam2',
      'component2','comp2','internal2','int2','secondtest','periodicaltest2',
      'pt2','periodical2','cia2'],
    cat3: ['cat3','cat3marks','ca3','ct3','ut3','test3','t3','exam3',
      'component3','comp3','internal3','int3','thirdtest','periodicaltest3',
      'pt3','periodical3','cia3'],
    assign1: ['assignment1','assign1','ass1','a1','asgn1','hw1','homework1',
      'project1','task1','lab1','practical1','prac1','record1'],
    assign2: ['assignment2','assign2','ass2','a2','asgn2','hw2','homework2',
      'project2','task2','lab2','practical2','prac2','record2'],
    quiz: ['quiz','quizmarks','quiztotal','viva','vivamarks','oral',
      'oralmarks','seminar','presentation'],
  };

  /* Reverse lookup: normalisedAlias → role */
  const REV = {};
  Object.entries(DICT).forEach(([role, aliases]) =>
    aliases.forEach(a => { REV[UTILS.norm(a)] = role; }));

  /* Regex patterns for structural header names */
  const PATTERNS = [
    { role:'cat1',    re:/(?:cat|component|test|exam|ct|ut|pt|cia)\s*[-_]?\s*1\b/i },
    { role:'cat2',    re:/(?:cat|component|test|exam|ct|ut|pt|cia)\s*[-_]?\s*2\b/i },
    { role:'cat3',    re:/(?:cat|component|test|exam|ct|ut|pt|cia)\s*[-_]?\s*3\b/i },
    { role:'assign1', re:/(?:assign|asgn|lab|prac|proj|record)\s*[-_]?\s*1\b/i },
    { role:'assign2', re:/(?:assign|asgn|lab|prac|proj|record)\s*[-_]?\s*2\b/i },
    { role:'internal',re:/int(?:ernal)?\s*(?:total|marks?|score|avg|assessment)?/i },
    { role:'external',re:/ext(?:ernal)?\s*(?:total|marks?|score|exam)?/i },
    { role:'total',   re:/tot(?:al)?\s*(?:marks?|score)?/i },
    { role:'gpa',     re:/\b[csg]?pa\b/i },
    { role:'roll',    re:/(?:roll|reg|enrol|admission|ht|hallticket)\s*(?:no\.?|num|id)?/i },
    { role:'grade',   re:/^gr(?:ade)?s?$/i },
    { role:'result',  re:/\b(?:result|status|passfail|outcome|remarks?)\b/i },
  ];

  /** Score a header against a role — higher = more confident */
  function score(header, role) {
    const n = UTILS.norm(header);
    if (REV[n] === role) return 100;           // exact alias
    if (DICT[role]?.some(a => n.includes(UTILS.norm(a)))) return 70; // contains
    if (PATTERNS.find(p => p.role === role && p.re.test(header))) return 60; // regex
    return 0;
  }

  /** Sample column values to infer max and numeric nature */
  function sampleCol(rows, col) {
    const vals = rows.slice(0, 120).map(r => UTILS.n(r[col])).filter(v => v !== null);
    if (!vals.length) return { numeric: false };
    const max  = Math.max(...vals);
    const mean = vals.reduce((s, v) => s + v, 0) / vals.length;
    const nzPct = vals.filter(v => v > 0).length / vals.length;
    return { numeric: true, max, mean, nzPct };
  }

  /** Snap a discovered max to a known academic milestone */
  function snapMax(m) {
    const ms = [5, 10, 15, 20, 25, 30, 40, 50, 60, 75, 80, 100, 150, 200, 300, 500];
    return ms.find(x => x >= m) ?? Math.ceil(m / 10) * 10;
  }

  /**
   * Main detection entry point.
   * Returns: { map, components, mode, maxes, issues }
   * map: role → { col, max, conf }
   */
  function detect(rows) {
    if (!rows.length) return { map:{}, components:[], mode:'bare', maxes:{}, issues:[] };

    const headers = Object.keys(rows[0]);
    const samples = {}; // header → sample stats
    headers.forEach(h => { samples[h] = sampleCol(rows, h); });

    const maxes = {}; // header → snapped max
    headers.forEach(h => { maxes[h] = samples[h].numeric ? snapMax(samples[h].max) : null; });

    /* Score every header against every role, take best per role */
    const scores = {}; // role → { col, sc }
    const multiRoles = new Set(['cat1','cat2','cat3','assign1','assign2','quiz']);

    headers.forEach(h => {
      const roles = Object.keys(DICT);
      roles.forEach(role => {
        const sc = score(h, role);
        if (sc === 0) return;
        if (!scores[role] || sc > scores[role].sc) {
          scores[role] = { col: h, sc, max: maxes[h] ?? 100 };
        }
      });
    });

    const map = {};
    Object.entries(scores).forEach(([role, v]) => {
      map[role] = { col: v.col, max: v.max ?? 100, conf: v.sc >= 100 ? 'exact' : v.sc >= 70 ? 'contains' : 'pattern' };
    });

    /* Extra instances of multi-roles (cat2 might not be detected if cat1 took it) */
    // Re-sweep: if a header's best role is already taken, assign as secondary
    const usedCols = new Set(Object.values(map).map(m => m.col));
    const extras = []; // { role, col, max }
    headers.forEach(h => {
      if (usedCols.has(h)) return;
      for (const { role, re } of PATTERNS) {
        if (re.test(h) && multiRoles.has(role)) {
          extras.push({ role, col: h, max: maxes[h] ?? 100 });
          usedCols.add(h);
          break;
        }
      }
    });

    /* Assemble components list (internal sub-scores) */
    const compRoles = ['cat1','cat2','cat3','assign1','assign2','quiz'];
    const components = [];
    compRoles.forEach(role => {
      if (map[role]) components.push({ role, col: map[role].col, max: map[role].max });
    });
    extras.forEach(e => {
      if (compRoles.includes(e.role)) components.push(e);
    });

    /* Fallbacks */
    if (!map.name) {
      const fc = headers.find(h => !samples[h].numeric);
      if (fc) map.name = { col: fc, max: null, conf: 'heuristic' };
    }
    if (!map.roll) {
      const sc = headers.filter(h => !samples[h].numeric)[1];
      if (sc && sc !== map.name?.col) map.roll = { col: sc, max: null, conf: 'heuristic' };
    }
    if (!map.total && !map.external) {
      const used = new Set(Object.values(map).map(m => m.col));
      const best = headers
        .filter(h => !used.has(h) && samples[h].numeric && samples[h].max >= 40)
        .sort((a, b) => (samples[b].mean ?? 0) - (samples[a].mean ?? 0))[0];
      if (best) map.external = { col: best, max: maxes[best] ?? 100, conf: 'heuristic' };
    }

    /* Mode */
    let mode;
    if (components.length >= 2)             mode = 'component';
    else if (map.internal && map.external)  mode = 'int_ext';
    else if (map.total)                     mode = 'total_only';
    else                                    mode = 'bare';

    /* Issues */
    const issues = [];
    if (!map.name && !map.roll) issues.push('No name or roll column found — student identification unavailable');
    if (!map.total && !map.external && !map.internal && !components.length)
      issues.push('No score column detected — cannot compute analytics');

    return { map, components, mode, maxes, issues };
  }

  return { detect, snapMax, score };
})();

/* ================================================================
   PARSER — SheetJS → plain row array
   ================================================================ */
const PARSER = {
  parse(arrayBuffer) {
    const wb = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws, { defval: '', raw: true });
  },
};

/* ================================================================
   NORMALIZER — raw rows → normalizedStudentData
   Every field has a stable name. All downstream code uses ONLY this.
   ================================================================ */
const NORMALIZER = {
  /**
   * Returns normalizedStudentData: Array<NormRow>
   * NormRow fields:
   *   _id, _name, _roll, _subject, _dept, _section
   *   _components: [{role,col,value,max,pct}]
   *   _internalRaw, _internalMax, _internalPct
   *   _externalRaw, _externalMax, _externalPct
   *   _totalRaw, _totalMax
   *   _pct          ← PRIMARY SCORE PERCENTAGE (0-100) — used everywhere
   *   _gpa, _grade, _pass (bool), _band, _rank (set later)
   *   _rawGrade, _rawResult (original strings for reference)
   *   _attendance
   *   _src          ← original row reference
   */
  normalize(rawRows, schema) {
    const { map, components, mode } = schema;
    const g = role => map[role]?.col ?? null;  // get column name for role

    // Step 1: remove completely empty rows
    const cleaned = rawRows.filter(row => {
      const vals = Object.values(row).map(v => String(v ?? '').trim()).filter(Boolean);
      return vals.length > 0;
    });

    // Step 2: build normalized rows
    const normalized = cleaned.map((row, idx) => {
      const nr = { _src: row, _id: idx };

      /* Identity */
      nr._name    = g('name')    ? String(row[g('name')] ?? '').trim() : '';
      nr._roll    = g('roll')    ? String(row[g('roll')] ?? '').trim() : '';
      nr._subject = g('subject') ? String(row[g('subject')] ?? '').trim() : '';
      nr._dept    = g('department') ? String(row[g('department')] ?? '').trim() : '';
      nr._section = g('section')    ? String(row[g('section')] ?? '').trim() : '';

      /* Components */
      nr._components = components.map(cp => {
        const raw = UTILS.n(row[cp.col]);
        const value = raw ?? 0;
        const max = cp.max || 100;
        return { role: cp.role, col: cp.col, value, max, pct: max > 0 ? UTILS.clamp((value / max) * 100, 100) : 0 };
      });

      /* Internal */
      let intRaw = null, intMax = 100;
      if (nr._components.length >= 2) {
        intRaw = nr._components.reduce((s, c) => s + c.value, 0);
        intMax = nr._components.reduce((s, c) => s + c.max, 0) || 100;
      } else if (g('internal')) {
        intRaw = UTILS.n(row[g('internal')]) ?? 0;
        intMax = map.internal?.max ?? 100;
      }
      nr._internalRaw = intRaw !== null ? UTILS.r(intRaw) : null;
      nr._internalMax = intMax;
      nr._internalPct = (intMax > 0 && intRaw !== null)
        ? UTILS.clamp(UTILS.r((intRaw / intMax) * 100), 100) : null;

      /* External */
      let extRaw = null, extMax = 100;
      if (g('external')) {
        extRaw = UTILS.n(row[g('external')]) ?? 0;
        extMax = map.external?.max ?? 100;
      }
      nr._externalRaw = extRaw !== null ? UTILS.r(extRaw) : null;
      nr._externalMax = extMax;
      nr._externalPct = (extMax > 0 && extRaw !== null)
        ? UTILS.clamp(UTILS.r((extRaw / extMax) * 100), 100) : null;

      /* Total */
      let totRaw = null, totMax = 100;
      if (g('total')) {
        totRaw = UTILS.n(row[g('total')]) ?? 0;
        totMax = map.total?.max ?? 100;
      }
      nr._totalRaw = totRaw !== null ? UTILS.r(totRaw) : null;
      nr._totalMax = totMax;

      /* Primary percentage */
      let pct = 0;
      if (mode === 'component') {
        const ip = nr._internalPct ?? 0;
        if (nr._externalPct !== null) {
          pct = ip * 0.40 + nr._externalPct * 0.60;
        } else { pct = ip; }
        if (totRaw !== null && totMax > 0) pct = (totRaw / totMax) * 100;
      } else if (mode === 'int_ext') {
        const ip = nr._internalPct ?? 0;
        const ep = nr._externalPct ?? 0;
        pct = ip * 0.40 + ep * 0.60;
        if (totRaw !== null && totMax > 0) pct = (totRaw / totMax) * 100;
      } else if (mode === 'total_only') {
        pct = totMax > 0 ? (totRaw / totMax) * 100 : 0;
      } else {
        // bare: use whatever numeric we have
        const first = [
          g('external') && extRaw !== null ? { v: extRaw, m: extMax } : null,
          g('internal') && intRaw !== null ? { v: intRaw, m: intMax } : null,
          g('total')    && totRaw !== null ? { v: totRaw, m: totMax } : null,
        ].filter(Boolean)[0];
        if (first && first.m > 0) pct = (first.v / first.m) * 100;
      }
      nr._pct = UTILS.clamp(UTILS.r(pct, 2), 100);

      /* GPA */
      if (g('gpa') && UTILS.n(row[g('gpa')]) !== null) {
        nr._gpa = UTILS.clamp(+UTILS.n(row[g('gpa')]).toFixed(2), 10);
      } else {
        nr._gpa = FAIL_DETECTOR.pctToGPA(nr._pct);
      }

      /* Raw strings for grade and result */
      nr._rawGrade  = g('grade')  ? String(row[g('grade')]  ?? '').trim() : '';
      nr._rawResult = g('result') ? String(row[g('result')] ?? '').trim() : '';
      nr._attendance= g('attendance') ? UTILS.n(row[g('attendance')]) : null;

      /* Grade */
      nr._grade = nr._rawGrade || FAIL_DETECTOR.gpaToGrade(nr._gpa);

      /* Pass/Fail — via SINGLE authoritative detector */
      nr._pass = FAIL_DETECTOR.isPass(nr);

      /* Band */
      nr._band = FAIL_DETECTOR.band(nr._pct);

      return nr;
    });

    // Step 3: remove rows with zero pct AND empty name/roll (likely header artifacts)
    const valid = normalized.filter(nr =>
      nr._pct > 0 || nr._name || nr._roll
    );

    // Step 4: remove duplicate rows (same name+roll+subject with identical score)
    const seen = new Set();
    const deduped = valid.filter(nr => {
      const key = `${nr._name}|${nr._roll}|${nr._subject}|${nr._pct}`;
      if (seen.has(key)) return false;
      seen.add(key); return true;
    });

    // Step 5: assign ranks
    const ranked = [...deduped].sort((a, b) => b._pct - a._pct);
    ranked.forEach((nr, i) => { nr._rank = i + 1; });

    return deduped;
  },
};

/* ================================================================
   FAIL_DETECTOR — ONE authoritative pass/fail source
   Every module MUST call FAIL_DETECTOR.isPass(nr)
   ================================================================ */
const FAIL_DETECTOR = {
  /** Strings that unambiguously mean FAIL */
  FAIL_STRINGS: new Set(['fail','f','ra','u','absent','ab','debarred','detained',
    'withheld','wh','xx','--','na','0']),
  /** Strings that mean PASS */
  PASS_STRINGS: new Set(['pass','p','passed','o','a+','a','b+','b','c','d','present',
    'admitted','promoted','yes','y','1','true','ok','clear','through']),

  /**
   * Master pass/fail detector.
   * Checks (in order of priority):
   *  1. Raw result string
   *  2. Raw grade string
   *  3. Grade field
   *  4. External percentage threshold
   *  5. Total percentage threshold
   */
  isPass(nr) {
    // 1. Raw result string
    const rr = UTILS.norm(nr._rawResult);
    if (rr && this.FAIL_STRINGS.has(rr)) return false;
    if (rr && this.PASS_STRINGS.has(rr)) return true;

    // 2. Raw grade string
    const rg = UTILS.norm(nr._rawGrade);
    if (rg && this.FAIL_STRINGS.has(rg)) return false;
    if (rg === 'f' || rg === 'fail' || rg === 'ra' || rg === 'u') return false;

    // 3. Computed grade
    const cg = UTILS.norm(nr._grade);
    if (cg === 'f') return false;

    // 4. External threshold: must score ≥35% in external to pass
    if (nr._externalPct !== null && nr._externalPct < 35) return false;

    // 5. Total percentage threshold
    if (nr._pct < 40) return false;

    return true;
  },

  pctToGPA(pct) {
    if (pct >= 91) return 10; if (pct >= 81) return 9;  if (pct >= 71) return 8;
    if (pct >= 61) return 7;  if (pct >= 51) return 6;  if (pct >= 41) return 5;
    if (pct >= 30) return 4;  return 0;
  },

  gpaToGrade(gpa) {
    if (gpa >= 10) return 'O';  if (gpa >= 9) return 'A+'; if (gpa >= 8) return 'A';
    if (gpa >= 7)  return 'B+'; if (gpa >= 6) return 'B';  if (gpa >= 5) return 'C';
    if (gpa >= 4)  return 'D';  return 'F';
  },

  band(pct) {
    if (pct >= 85) return 'Excellent';
    if (pct >= 70) return 'Good';
    if (pct >= 50) return 'Average';
    if (pct >= 40) return 'Weak';
    return 'Failed';
  },

  bandColor(band) {
    return { Excellent:'#facc15', Good:'#86efac', Average:'#fdba74',
             Weak:'#fca5a5', Failed:'#ef4444' }[band] ?? '#9ca3af';
  },
};

/* ================================================================
   ANALYTICS — pure reusable functions, all take normalizedStudentData
   ================================================================ */
const ANALYTICS = {
  /* ── Descriptive ── */
  avg(arr) {
    const v = arr.filter(x => x !== null && !isNaN(x));
    return v.length ? v.reduce((s, x) => s + x, 0) / v.length : 0;
  },
  median(arr) {
    const v = [...arr.filter(x => x !== null && !isNaN(x))].sort((a, b) => a - b);
    if (!v.length) return 0;
    const m = Math.floor(v.length / 2);
    return v.length % 2 ? v[m] : (v[m-1] + v[m]) / 2;
  },
  stdDev(arr) {
    const v = arr.filter(x => x !== null && !isNaN(x));
    if (v.length < 2) return 0;
    const mean = this.avg(v);
    return Math.sqrt(v.reduce((s, x) => s + (x - mean) ** 2, 0) / v.length);
  },
  percentile(arr, p) {
    const v = [...arr.filter(x => x !== null && !isNaN(x))].sort((a, b) => a - b);
    if (!v.length) return 0;
    const i = (p / 100) * (v.length - 1);
    const lo = Math.floor(i), hi = Math.ceil(i);
    return v[lo] + (v[hi] - v[lo]) * (i - lo);
  },
  max(arr) { return arr.length ? Math.max(...arr.filter(x => x !== null && !isNaN(x))) : 0; },
  min(arr) { return arr.length ? Math.min(...arr.filter(x => x !== null && !isNaN(x))) : 0; },

  /* ── Academic analytics on normalizedStudentData ── */

  /** Core stats object — single source for all KPIs */
  compute(data) {
    if (!data.length) return this._empty();
    const pcts  = data.map(r => r._pct);
    const gpas  = data.map(r => r._gpa);
    const passes = data.map(r => r._pass);

    const passCount = passes.filter(Boolean).length;
    const failCount = data.length - passCount;

    const bandCounts = { Excellent:0, Good:0, Average:0, Weak:0, Failed:0 };
    data.forEach(r => bandCounts[r._band]++);

    const gradeDist = this.gradeDist(data.map(r => r._grade));
    const scoreBands = this.scoreBands(pcts);
    const gpaStats = this.gpaStats(gpas);
    const subjectStats = this.groupAnalysis(data, r => r._subject || null);
    const deptStats    = this.groupAnalysis(data, r => r._dept    || null);
    const sectionStats = this.groupAnalysis(data, r => r._section || null);

    const compAvgs = {};
    APP.schema.components.forEach(cp => {
      const vals = data.map(r => {
        const c = r._components.find(x => x.col === cp.col);
        return c ? c.pct : null;
      }).filter(v => v !== null);
      compAvgs[cp.col] = { label: cp.col, role: cp.role, avg: UTILS.r(this.avg(vals)) };
    });

    const intExtCorr = (data[0]?._internalPct !== null && data[0]?._externalPct !== null)
      ? this.pearson(data.map(r => r._internalPct ?? 0), data.map(r => r._externalPct ?? 0))
      : null;

    const uniq = fn => [...new Set(data.map(fn).filter(Boolean))];

    return {
      total: data.length, passCount, failCount,
      passRate: UTILS.r((passCount / data.length) * 100),
      failRate: UTILS.r((failCount / data.length) * 100),
      avg: UTILS.r(this.avg(pcts)),
      median: UTILS.r(this.median(pcts)),
      stdDev: UTILS.r(this.stdDev(pcts)),
      p25: UTILS.r(this.percentile(pcts, 25)),
      p75: UTILS.r(this.percentile(pcts, 75)),
      high: UTILS.r(this.max(pcts)),
      low: UTILS.r(this.min(pcts)),
      topper: data.reduce((best, r) => (!best || r._pct > best._pct ? r : best), null),
      bandCounts, gradeDist, scoreBands, gpaStats,
      subjectStats, deptStats, sectionStats,
      compAvgs, intExtCorr,
      avgGPA: gpaStats.avg,
      weakCount: bandCounts.Weak + bandCounts.Failed,
      subjects:    uniq(r => r._subject),
      departments: uniq(r => r._dept),
      sections:    uniq(r => r._section),
      grades:      Object.keys(gradeDist),
    };
  },

  _empty() {
    return { total:0, passCount:0, failCount:0, passRate:0, failRate:0,
      avg:0, median:0, stdDev:0, high:0, low:0, topper:null,
      bandCounts:{Excellent:0,Good:0,Average:0,Weak:0,Failed:0},
      gradeDist:{}, scoreBands:{}, gpaStats:this.gpaStats([]),
      subjectStats:[], deptStats:[], sectionStats:[], compAvgs:{},
      intExtCorr:null, avgGPA:0, weakCount:0,
      subjects:[], departments:[], sections:[], grades:[], p25:0, p75:0 };
  },

  /** Grade distribution — ordered */
  gradeDist(grades) {
    const dist = {};
    const order = ['O','A+','A','B+','B','C','D','F'];
    grades.forEach(g => { if (g) dist[g] = (dist[g] || 0) + 1; });
    const sorted = {};
    [...order, ...Object.keys(dist).filter(k => !order.includes(k))]
      .forEach(k => { if (dist[k]) sorted[k] = dist[k]; });
    return sorted;
  },

  /** Score band histogram */
  scoreBands(pcts) {
    const b = {'0–39':0,'40–49':0,'50–59':0,'60–69':0,'70–79':0,'80–89':0,'90–100':0};
    pcts.forEach(p => {
      if      (p < 40) b['0–39']++;   else if (p < 50) b['40–49']++;
      else if (p < 60) b['50–59']++;  else if (p < 70) b['60–69']++;
      else if (p < 80) b['70–79']++;  else if (p < 90) b['80–89']++;
      else             b['90–100']++;
    });
    return b;
  },

  /** GPA stats + distribution */
  gpaStats(gpas) {
    const v = gpas.filter(x => x !== null && !isNaN(x));
    const buckets = {'0':0,'4':0,'5':0,'6':0,'7':0,'8':0,'9':0,'10':0};
    v.forEach(g => {
      const k = String(Math.floor(g));
      buckets.hasOwnProperty(k) ? buckets[k]++ : buckets['0']++;
    });
    return {
      avg: UTILS.r(this.avg(v), 2), median: UTILS.r(this.median(v), 2),
      distribution: buckets,
      excellent: v.filter(g => g >= 9).length,
      aboveAvg:  v.filter(g => g >= 7 && g < 9).length,
      average:   v.filter(g => g >= 5 && g < 7).length,
      below:     v.filter(g => g > 0 && g < 5).length,
      failed:    v.filter(g => g === 0).length,
    };
  },

  /** Group analysis — pass getLabel to group by subject/dept/section */
  groupAnalysis(data, getLabel) {
    const groups = {};
    data.forEach(r => {
      const lbl = getLabel(r); if (!lbl) return;
      if (!groups[lbl]) groups[lbl] = { pcts:[], pass:0, fail:0 };
      groups[lbl].pcts.push(r._pct);
      r._pass ? groups[lbl].pass++ : groups[lbl].fail++;
    });
    return Object.entries(groups).map(([label, g]) => ({
      label,
      avg:       UTILS.r(this.avg(g.pcts)),
      median:    UTILS.r(this.median(g.pcts)),
      max:       UTILS.r(this.max(g.pcts)),
      passCount: g.pass, failCount: g.fail,
      total:     g.pcts.length,
      passRate:  g.pcts.length ? UTILS.r((g.pass / g.pcts.length) * 100) : 0,
      failRate:  g.pcts.length ? UTILS.r((g.fail / g.pcts.length) * 100) : 0,
    })).sort((a, b) => b.avg - a.avg);
  },

  /** Pearson correlation */
  pearson(xs, ys) {
    const n = Math.min(xs.length, ys.length);
    if (n < 3) return null;
    const mx = this.avg(xs.slice(0,n)), my = this.avg(ys.slice(0,n));
    let num = 0, dx2 = 0, dy2 = 0;
    for (let i = 0; i < n; i++) {
      const dx = xs[i] - mx, dy = ys[i] - my;
      num += dx * dy; dx2 += dx*dx; dy2 += dy*dy;
    }
    return dx2 && dy2 ? +(num / Math.sqrt(dx2 * dy2)).toFixed(3) : null;
  },

  /** Fail percentage */
  failPct(data) {
    if (!data.length) return 0;
    return UTILS.r((data.filter(r => !r._pass).length / data.length) * 100);
  },

  /** Pass percentage */
  passPct(data) { return UTILS.r(100 - this.failPct(data)); },

  /** Rankings — returns sorted by _pct desc */
  rankings(data) {
    return [...data].sort((a, b) => b._pct - a._pct).map((r, i) => ({ ...r, _rank: i+1 }));
  },
};

/* ================================================================
   CHART_ENGINE — Chart.js wrappers, all data from ANALYTICS
   ================================================================ */
const CHART_ENGINE = (() => {
  const inst = {};

  const YELLOWS = [
    'rgba(250,204,21,.82)','rgba(234,179,8,.82)','rgba(253,224,71,.82)',
    'rgba(217,119,6,.85)','rgba(251,191,36,.82)','rgba(161,98,7,.82)',
    'rgba(245,158,11,.82)','rgba(252,211,77,.82)',
  ];
  const MULTI = [
    'rgba(250,204,21,.82)','rgba(34,197,94,.78)','rgba(239,68,68,.78)',
    'rgba(59,130,246,.78)','rgba(168,85,247,.78)','rgba(249,115,22,.78)',
    'rgba(20,184,166,.78)','rgba(236,72,153,.78)','rgba(132,204,22,.78)',
  ];

  function base(noScales = false) {
    const cfg = {
      responsive: true, maintainAspectRatio: false,
      animation: { duration: 650, easing: 'easeOutQuart' },
      plugins: {
        legend: { labels: { color:'#6b7280', font:{ family:'Inter', size:11.5 }, usePointStyle:true, padding:14 } },
        tooltip: {
          backgroundColor:'rgba(255,255,255,.97)', borderColor:'rgba(250,204,21,.35)',
          borderWidth:1, titleColor:'#111827', bodyColor:'#6b7280',
          titleFont:{ family:'Space Grotesk', size:13 },
          bodyFont:{ family:'JetBrains Mono', size:11.5 }, padding:12,
        },
      },
    };
    if (!noScales) cfg.scales = {
      x:{ ticks:{ color:'#9ca3af', font:{ family:'JetBrains Mono', size:10.5 }, maxRotation:38 }, grid:{ color:'rgba(0,0,0,.04)' } },
      y:{ beginAtZero:true, ticks:{ color:'#9ca3af', font:{ family:'JetBrains Mono', size:10.5 } }, grid:{ color:'rgba(0,0,0,.05)' } },
    };
    return cfg;
  }

  function mk(key, id, config) {
    if (inst[key]) { inst[key].destroy(); delete inst[key]; }
    const c = UTILS.$(id); if (!c) return;
    inst[key] = new Chart(c.getContext('2d'), config);
  }

  /* ── Individual chart builders ── */

  function subjectAvg(S, type = 'bar') {
    const data = S.subjectStats.slice(0, 16);
    if (!data.length) { fallbackBar('subjectAvg','cSubjectAvg', S); return; }
    const cfg = base(); cfg.plugins.legend.display = false; cfg.scales.y.max = 100;
    mk('subjectAvg','cSubjectAvg', { type,
      data:{ labels: data.map(d => d.label),
        datasets:[{ label:'Avg %', data: data.map(d => d.avg),
          backgroundColor: data.map((_,i) => YELLOWS[i%YELLOWS.length]),
          borderColor: data.map((_,i) => YELLOWS[i%YELLOWS.length].replace('.82','1')),
          borderWidth: type === 'line' ? 2 : 0,
          borderRadius: type === 'bar' ? 8 : 0,
          tension:.4, fill: type === 'line',
          pointBackgroundColor:'rgba(250,204,21,1)' }] }, options: cfg });
  }

  function fallbackBar(key, id, S) {
    const data = APP.data.slice(0, 30);
    const cfg = base(); cfg.plugins.legend.display = false; cfg.scales.y.max = 100;
    mk(key, id, { type:'bar',
      data:{ labels: data.map((r,i) => r._name || r._roll || `S${i+1}`),
        datasets:[{ data: data.map(r => r._pct), backgroundColor:'rgba(250,204,21,.75)',
          borderRadius:6, borderWidth:0 }] }, options:cfg });
  }

  function subjectPF(S) {
    const data = S.subjectStats.slice(0, 12);
    if (!data.length) return;
    const cfg = base(); cfg.scales.x.stacked = true; cfg.scales.y.stacked = true;
    mk('subjectPF','cSubjectPF', { type:'bar',
      data:{ labels: data.map(d => d.label),
        datasets:[
          { label:'Pass', data: data.map(d => d.passCount), backgroundColor:'rgba(34,197,94,.75)', borderRadius:6 },
          { label:'Fail', data: data.map(d => d.failCount), backgroundColor:'rgba(239,68,68,.7)',  borderRadius:6 },
        ] }, options:cfg });
  }

  function passFail(S) {
    mk('passFail','cPassFail', { type:'doughnut',
      data:{ labels:['Pass','Fail'],
        datasets:[{ data:[S.passCount, S.failCount],
          backgroundColor:['rgba(34,197,94,.82)','rgba(239,68,68,.78)'],
          borderColor:['rgba(255,255,255,.9)','rgba(255,255,255,.9)'],
          borderWidth:3, hoverOffset:8 }] },
      options:{ ...base(true), cutout:'64%' } });
  }

  function gradeChart(S) {
    const entries = Object.entries(S.gradeDist); if (!entries.length) return;
    mk('grade','cGrade', { type:'doughnut',
      data:{ labels:entries.map(([k])=>k),
        datasets:[{ data:entries.map(([,v])=>v), backgroundColor:MULTI,
          borderColor:'rgba(255,255,255,.85)', borderWidth:2, hoverOffset:8 }] },
      options:{ ...base(true), cutout:'55%' } });
  }

  function scoreBands(S) {
    const e = Object.entries(S.scoreBands);
    const cfg = base(); cfg.plugins.legend.display = false;
    mk('scoreBands','cScoreBands', { type:'bar',
      data:{ labels:e.map(([k])=>k),
        datasets:[{ label:'Students', data:e.map(([,v])=>v),
          backgroundColor:e.map(([k])=>{
            const p=parseInt(k);
            return p>=90?'rgba(250,204,21,.85)':p>=70?'rgba(134,239,172,.8)':p>=50?'rgba(253,186,116,.8)':p>=40?'rgba(252,165,165,.8)':'rgba(239,68,68,.75)';
          }), borderRadius:8, borderWidth:0 }] }, options:cfg });
  }

  function gpaChart(S) {
    const e = Object.entries(S.gpaStats.distribution);
    const cfg = base(); cfg.plugins.legend.display = false;
    mk('gpa','cGPA', { type:'bar',
      data:{ labels:e.map(([k])=>k),
        datasets:[{ label:'Students', data:e.map(([,v])=>v),
          backgroundColor:'rgba(250,204,21,.75)',
          borderColor:'rgba(234,179,8,1)', borderWidth:1, borderRadius:8 }] }, options:cfg });
  }

  function bandPie(S) {
    const entries = Object.entries(S.bandCounts).filter(([,v])=>v>0);
    mk('bandPie','cBandPie', { type:'pie',
      data:{ labels:entries.map(([k])=>k),
        datasets:[{ data:entries.map(([,v])=>v),
          backgroundColor:entries.map(([k])=>FAIL_DETECTOR.bandColor(k)+'cc'),
          borderColor:'rgba(255,255,255,.9)', borderWidth:2, hoverOffset:6 }] },
      options: base(true) });
  }

  function intVsExt() {
    const valid = APP.data.filter(r => r._internalPct !== null && r._externalPct !== null).slice(0,60);
    if (!valid.length) return;
    const cfg = base(); cfg.elements = { point:{radius:3}, line:{tension:.4} }; cfg.scales.y.max = 100;
    mk('intVsExt','cIntVsExt', { type:'line',
      data:{ labels:valid.map((r,i)=>(r._name||r._roll||`S${i+1}`).slice(0,10)),
        datasets:[
          { label:'Internal %', data:valid.map(r=>r._internalPct), borderColor:'rgba(250,204,21,1)', backgroundColor:'rgba(250,204,21,.08)', fill:true, borderWidth:2.5, pointBackgroundColor:'rgba(250,204,21,.9)' },
          { label:'External %', data:valid.map(r=>r._externalPct), borderColor:'rgba(17,24,39,.7)',  backgroundColor:'rgba(17,24,39,.04)',   fill:true, borderWidth:2,   pointBackgroundColor:'rgba(17,24,39,.7)'  },
        ] }, options:cfg });
  }

  function components(S) {
    const comps = Object.values(S.compAvgs); if (!comps.length) return;
    const cfg = base(); cfg.plugins.legend.display = false; cfg.scales.y.max = 100;
    mk('components','cComponents', { type:'bar',
      data:{ labels:comps.map(c=>c.label),
        datasets:[{ label:'Avg %', data:comps.map(c=>c.avg),
          backgroundColor:comps.map((_,i)=>YELLOWS[i%YELLOWS.length]),
          borderRadius:8, borderWidth:0 }] }, options:cfg });
  }

  function dept(S) {
    const data = S.deptStats.slice(0,12); if (!data.length) return;
    const cfg = base(); cfg.indexAxis='y'; cfg.plugins.legend.display=false; cfg.scales.x.max=100;
    mk('dept','cDept', { type:'bar',
      data:{ labels:data.map(d=>d.label),
        datasets:[{ label:'Avg %', data:data.map(d=>d.avg),
          backgroundColor:data.map((_,i)=>YELLOWS[i%YELLOWS.length]),
          borderRadius:8, borderWidth:0 }] }, options:cfg });
  }

  function trend() {
    const sorted = [...APP.data].sort((a,b)=>(a._rank||0)-(b._rank||0)).slice(0,60);
    const cfg = base(); cfg.elements={point:{radius:4},line:{tension:.45}}; cfg.plugins.legend.display=false; cfg.scales.y.max=100;
    mk('trend','cTrend', { type:'line',
      data:{ labels:sorted.map((r,i)=>(r._name||r._roll||`#${i+1}`).slice(0,10)),
        datasets:[{ label:'Score %', data:sorted.map(r=>r._pct),
          borderColor:'rgba(250,204,21,1)', backgroundColor:'rgba(250,204,21,.07)',
          fill:true, borderWidth:2.5,
          pointBackgroundColor:sorted.map(r=>r._pass?'rgba(34,197,94,.85)':'rgba(239,68,68,.85)'),
          pointBorderColor:'#fff', pointBorderWidth:1.5, pointRadius:4 }] }, options:cfg });
  }

  function buildAll(S) {
    subjectAvg(S, 'bar');
    subjectPF(S);
    passFail(S);
    gradeChart(S);
    scoreBands(S);
    gpaChart(S);
    bandPie(S);
    intVsExt();
    components(S);
    dept(S);
    trend();
  }

  return { buildAll, subjectAvg, inst };
})();

/* ================================================================
   INSIGHTS_ENGINE — all text generated from ANALYTICS.compute()
   ================================================================ */
const INSIGHTS_ENGINE = {
  generate(data, S) {
    const ins = [];
    const f = v => typeof v === 'number' ? v.toFixed(1) : String(v ?? '—');
    const n = UTILS.n;

    // 1. Overall performance
    const perf = S.avg >= 75 ? 'strong' : S.avg >= 55 ? 'moderate' : 'below average';
    ins.push({ ico:'◈', title:'Overall Class Performance',
      text:`Class average <strong>${f(S.avg)}%</strong> (median ${f(S.median)}%) — <strong>${perf}</strong>. SD ${f(S.stdDev)} → ${S.stdDev<10?'uniform class':S.stdDev<20?'moderate spread':'high variance'}. IQR: ${f(S.p25)}–${f(S.p75)}%.` });

    // 2. Pass/Fail — uses ANALYTICS counts which use FAIL_DETECTOR
    const fl = S.failRate > 30 ? 'alarming' : S.failRate > 15 ? 'concerning' : 'acceptable';
    ins.push({ ico:'◉', title:'Pass / Fail Analysis',
      text:`<strong>${f(S.passRate)}% passed</strong> (${S.passCount}/${S.total}). Fail rate <strong>${f(S.failRate)}%</strong> — ${fl}. ${S.failCount} student${S.failCount!==1?'s':''} require academic intervention.` });

    // 3. Topper
    if (S.topper) {
      const nm = S.topper._name || S.topper._roll || 'Top student';
      ins.push({ ico:'▲', title:'Class Topper',
        text:`<strong>${nm}</strong> leads with <strong>${f(S.topper._pct)}%</strong>, GPA ${S.topper._gpa}, Grade ${S.topper._grade}, Rank #${S.topper._rank}.` });
    }

    // 4. Subject analysis
    if (S.subjectStats.length >= 2) {
      const best = S.subjectStats[0], worst = S.subjectStats.at(-1);
      ins.push({ ico:'▦', title:'Subject Analysis',
        text:`Best: <strong>${best.label}</strong> (avg ${f(best.avg)}%, fail ${f(best.failRate)}%). Hardest: <strong>${worst.label}</strong> (avg ${f(worst.avg)}%, fail ${f(worst.failRate)}%). Gap: ${f(best.avg-worst.avg)} pts.` });
    }

    // 5. Score concentration
    const topBand = Object.entries(S.scoreBands).sort((a,b)=>b[1]-a[1])[0];
    if (topBand && topBand[1] > 0) {
      const pct = f((topBand[1]/S.total)*100);
      ins.push({ ico:'◒', title:'Score Concentration',
        text:`<strong>${pct}% of students</strong> scored in the <strong>${topBand[0]}</strong> band (${topBand[1]} students) — the dominant performance tier.` });
    }

    // 6. Component performance
    const comps = Object.values(S.compAvgs);
    if (comps.length >= 2) {
      const sorted = [...comps].sort((a,b)=>b.avg-a.avg);
      ins.push({ ico:'◐', title:'Component Performance',
        text:`Strongest component: <strong>${sorted[0].label}</strong> (${f(sorted[0].avg)}%). Weakest: <strong>${sorted.at(-1).label}</strong> (${f(sorted.at(-1).avg)}%). Consider remediation for weaker components.` });
    }

    // 7. Internal–External correlation
    if (S.intExtCorr !== null) {
      const c = S.intExtCorr;
      const cl = Math.abs(c)>.7?'strong':Math.abs(c)>.4?'moderate':'weak';
      const cd = c>=0?'positive':'negative';
      ins.push({ ico:'✦', title:'Internal–External Correlation',
        text:`Pearson r = <strong>${c}</strong> — a <strong>${cl} ${cd}</strong> relationship. ${Math.abs(c)>.6?'Internal assessments reliably predict exam performance.':'Divergence detected — review assessment alignment.'}` });
    }

    // 8. GPA summary
    const G = S.gpaStats;
    ins.push({ ico:'✦', title:'GPA Summary',
      text:`Avg GPA <strong>${G.avg}</strong> (median ${G.median}). Excellence (≥9): ${G.excellent} • Above avg: ${G.aboveAvg} • Average: ${G.average} • Below: ${G.below} • Failed: ${G.failed}.` });

    // 9. At-risk
    if (S.weakCount > 0) {
      const failRows = data.filter(r => !r._pass);
      const avgGap = failRows.length
        ? f(40 - ANALYTICS.avg(failRows.map(r => r._pct))) : '0';
      ins.push({ ico:'◌', title:'At-Risk Students',
        text:`<strong>${S.weakCount} (${f((S.weakCount/S.total)*100)}%)</strong> in Weak/Failed bands. Average gap to pass threshold: <strong>${avgGap} pts</strong>. Targeted remediation recommended.` });
    }

    // 10. Department
    if (S.deptStats.length > 1) {
      const best = S.deptStats[0], worst = S.deptStats.at(-1);
      ins.push({ ico:'◈', title:'Department Comparison',
        text:`<strong>${best.label}</strong> leads at avg <strong>${f(best.avg)}%</strong>, ${f(best.passRate)}% pass. <strong>${worst.label}</strong> has lowest avg ${f(worst.avg)}%. Gap: ${f(best.avg-worst.avg)} pts.` });
    }

    // 11. High achievers
    const above75 = data.filter(r => r._pct >= 75).length;
    ins.push({ ico:'▲', title:'High Achievers',
      text:`<strong>${above75} (${f((above75/data.length)*100)}%)</strong> scored ≥75%. <strong>${S.bandCounts.Excellent}</strong> achieved Excellent (≥85%). Consider enrichment programs.` });

    // 12. Score spread
    ins.push({ ico:'▦', title:'Score Spread',
      text:`Range: <strong>${f(S.low)}%</strong> to <strong>${f(S.high)}%</strong> (${f(S.high-S.low)} pts). IQR ${f(S.p25)}–${f(S.p75)}% — ${(S.p75-S.p25)<20?'tight distribution':'wide spread'}.` });

    return ins;
  },
};

/* ================================================================
   APP — global state + orchestrator
   ================================================================ */
const APP = {
  data:     [],    // normalizedStudentData — THE single source of truth
  filtered: [],    // current filtered view (same structure)
  stats:    {},    // ANALYTICS.compute(filtered)
  schema:   {},    // COLUMN_MAP output
  tablePage:      1,
  tablePageSize:  10,
  sortCol:        null,
  sortDir:        'asc',
  tblSearchQ:     '',
  modalChart:     null,

  /** Full pipeline */
  run(rawRows, fileName, fileSize) {
    UTILS.loader('Detecting columns…');
    setTimeout(() => {
      // 1. Detect schema
      this.schema = COLUMN_MAP.detect(rawRows);
      UTILS.loader('Normalizing data…');
      setTimeout(() => {
        // 2. Normalize → single dataset
        this.data = NORMALIZER.normalize(rawRows, this.schema);
        this.filtered = [...this.data];

        UTILS.loader('Computing analytics…');
        setTimeout(() => {
          // 3. Compute stats
          this.stats = ANALYTICS.compute(this.filtered);

          UTILS.loader('Building charts…');
          setTimeout(() => {
            // 4. Render everything
            UI.renderSchema(this.schema);
            UI.renderKPIs(this.stats);
            UI.renderBands(this.stats);
            UI.renderHeroCard(this.stats);
            UI.populateFilters(this.stats);
            CHART_ENGINE.buildAll(this.stats);
            UI.renderTable();
            UI.renderInsights(INSIGHTS_ENGINE.generate(this.filtered, this.stats));

            // Show success
            UTILS.$('uploadIdle').classList.add('hidden');
            UTILS.show('uploadDone');
            UTILS.$('uploadFileName').textContent = fileName;
            UTILS.$('uploadFileInfo').textContent =
              `${this.data.length} rows (${rawRows.length} raw) • ${Object.keys(rawRows[0]).length} cols • ${(fileSize/1024).toFixed(1)} KB • mode: ${this.schema.mode}`;
            UTILS.show('schemaPanel');
            UTILS.$('hcardDot')?.classList.add('live');
            UTILS.loaderOff();
            UTILS.toast(`${this.data.length} records analysed — schema: ${this.schema.mode}`);
          }, 50);
        }, 20);
      }, 20);
    }, 20);
  },

  applyFilters() {
    const sub  = UTILS.$('fSubject')?.value;
    const dept = UTILS.$('fDept')?.value;
    const sec  = UTILS.$('fSection')?.value;
    const grd  = UTILS.$('fGrade')?.value;
    const st   = UTILS.$('fStatus')?.value;

    this.filtered = this.data.filter(r => {
      if (sub  && r._subject !== sub)  return false;
      if (dept && r._dept    !== dept) return false;
      if (sec  && r._section !== sec)  return false;
      if (grd  && r._grade   !== grd)  return false;
      if (st === 'pass' && !r._pass)  return false;
      if (st === 'fail' &&  r._pass)  return false;
      if (['excellent','good','average','weak'].includes(st) && r._band.toLowerCase() !== st) return false;
      return true;
    });

    this.stats = ANALYTICS.compute(this.filtered);
    this.tablePage = 1;
    UI.renderKPIs(this.stats);
    UI.renderBands(this.stats);
    CHART_ENGINE.buildAll(this.stats);
    UI.renderTable();
    UI.renderInsights(INSIGHTS_ENGINE.generate(this.filtered, this.stats));
  },
};

/* ================================================================
   UI — DOM manipulation layer
   ================================================================ */
const UI = {
  renderSchema(schema) {
    const { map, components, mode, issues } = schema;
    const el = UTILS.$('schemaMode');
    if (el) el.textContent = mode;

    const roleLabels = {
      name:'Name', roll:'Roll/ID', subject:'Subject', department:'Department',
      section:'Section', internal:'Internal', external:'External', total:'Total',
      grade:'Grade', gpa:'GPA', result:'Result', attendance:'Attendance',
      cat1:'CAT1', cat2:'CAT2', cat3:'CAT3', assign1:'Assign1', assign2:'Assign2', quiz:'Quiz',
    };

    // Mode chip first
    const modeChip = `<div class="chip found" style="grid-column:1/-1">
      <span class="chip-k">Schema</span>
      <span class="chip-v"><strong>${mode}</strong> — ${components.length} component(s), intMax=${components.reduce((s,c)=>s+c.max,0)||'N/A'}</span>
    </div>`;

    const chips = Object.entries(roleLabels).map(([role, label]) => {
      const d = map[role];
      return `<div class="chip ${d?'found':'missing'}">
        <span class="chip-k">${label}</span>
        <span class="chip-v ${d?'':'miss'}">${d
          ? `${d.col} <em style="opacity:.55;font-size:9.5px">[${d.conf}${d.max!=null?', max='+d.max:''}]</em>`
          : '— not found'}</span>
      </div>`;
    }).join('');

    const chipsEl = UTILS.$('schemaChips');
    if (chipsEl) chipsEl.innerHTML = modeChip + chips;

    const issuesEl = UTILS.$('schemaIssues');
    if (issuesEl) issuesEl.innerHTML = issues.map(i => `⚠ ${i}`).join('<br>');
  },

  renderKPIs(S) {
    UTILS.animCount('kpiTotal', S.total);
    const kpiPass = UTILS.$('kpiPass'); if (kpiPass) kpiPass.textContent = S.passRate + '%';
    UTILS.animCount('kpiFail', S.failCount);
    const kpiHigh = UTILS.$('kpiHigh'); if (kpiHigh) kpiHigh.textContent = S.high + '%';
    const kpiAvg  = UTILS.$('kpiAvg');  if (kpiAvg)  kpiAvg.textContent  = S.avg  + '%';
    const kpiGPA  = UTILS.$('kpiGPA');  if (kpiGPA)  kpiGPA.textContent  = S.gpaStats?.avg ?? '—';
    UTILS.animCount('kpiSubjects', S.subjects.length);
    UTILS.animCount('kpiWeak', S.weakCount);
  },

  renderBands(S) {
    const B = S.bandCounts; const total = S.total || 1;
    ['Excellent','Good','Average','Weak','Failed'].forEach(name => {
      const c = B[name] || 0;
      UTILS.animCount('b' + name, c);
      const bar = UTILS.$('b' + name + 'Fill');
      if (bar) setTimeout(() => bar.style.width = ((c/total)*100).toFixed(1) + '%', 200);
    });
  },

  renderHeroCard(S) {
    const set = (id, v) => { const el = UTILS.$(id); if (el) el.textContent = v; };
    set('hcTotal', S.total);
    set('hcPass',  S.passRate + '%');
    set('hcAvg',   S.avg + '%');
    const bar = (bId, lId, pct, label) => {
      const b = UTILS.$(bId); if (b) setTimeout(() => b.style.width = Math.min(pct,100) + '%', 350);
      const l = UTILS.$(lId); if (l) l.textContent = label;
    };
    bar('hcBar1','hcBar1V', S.passRate, S.passRate + '%');
    bar('hcBar2','hcBar2V', S.avg, S.avg + '%');
    const exPct = S.bandCounts.Excellent
      ? ((S.bandCounts.Excellent / S.total) * 100).toFixed(0) : 0;
    bar('hcBar3','hcBar3V', exPct, exPct + '%');
  },

  populateFilters(S) {
    const fill = (id, opts, ph) => {
      const el = UTILS.$(id); if (!el) return;
      el.innerHTML = `<option value="">${ph}</option>` +
        opts.map(o => `<option value="${o}">${o}</option>`).join('');
    };
    fill('fSubject', S.subjects,    'All Subjects');
    fill('fDept',    S.departments, 'All Departments');
    fill('fSection', S.sections,    'All Sections');
    fill('fGrade',   S.grades,      'All Grades');
  },

  renderInsights(list) {
    const el = UTILS.$('insightsGrid'); if (!el) return;
    el.innerHTML = list.map((ins, i) => `
      <div class="insight-card glass-card" style="animation-delay:${i*.065}s">
        <div class="ins-ico">${ins.ico}</div>
        <div>
          <div class="ins-title">${ins.title}</div>
          <div class="ins-text">${ins.text}</div>
        </div>
      </div>`).join('');
  },

  /* ── Table ── */
  renderTable() {
    APP.tblSearchQ = '';
    const ts = UTILS.$('tblSearch'); if (ts) ts.value = '';
    APP.tablePage = 1;
    this._renderPage();
  },

  _getRows() {
    let rows = APP.filtered;
    if (APP.tblSearchQ) {
      const q = APP.tblSearchQ;
      rows = rows.filter(r =>
        [r._name, r._roll, r._subject, r._dept]
          .some(v => String(v ?? '').toLowerCase().includes(q))
        || String(r._grade ?? '').toLowerCase().includes(q)
        || (r._pass ? 'pass' : 'fail').includes(q)
      );
    }
    if (APP.sortCol) {
      rows = [...rows].sort((a, b) => {
        const va = a[APP.sortCol] ?? '', vb = b[APP.sortCol] ?? '';
        const na = parseFloat(va), nb = parseFloat(vb);
        if (!isNaN(na) && !isNaN(nb)) return APP.sortDir === 'asc' ? na-nb : nb-na;
        return APP.sortDir === 'asc' ? String(va).localeCompare(String(vb)) : String(vb).localeCompare(String(va));
      });
    }
    return rows;
  },

  _renderPage() {
    const rows = this._getRows();
    const cnt = UTILS.$('tblCount'); if (cnt) cnt.textContent = `${rows.length} records`;

    const wrap = UTILS.$('tblWrap');
    if (!rows.length) {
      if (wrap) wrap.innerHTML = '<p class="ph-msg">No records match your criteria.</p>';
      const pg = UTILS.$('pager'); if (pg) pg.innerHTML = '';
      return;
    }

    const schema = APP.schema;
    const map = schema.map || {};
    const totalPages = Math.ceil(rows.length / APP.tablePageSize);
    const start = (APP.tablePage - 1) * APP.tablePageSize;
    const pageRows = rows.slice(start, start + APP.tablePageSize);

    // Build dynamic columns from what was detected
    const cols = [
      { key:'_rank',    label:'#' },
      map.name    && { key:'_name',    label:'Name' },
      map.roll    && { key:'_roll',    label:'Roll/ID' },
      map.subject && { key:'_subject', label:'Subject' },
      map.department && { key:'_dept', label:'Dept' },
      map.section && { key:'_section', label:'Section' },
      ...(schema.components || []).map(cp => ({ key:`__comp_${cp.col}`, label:`${cp.role.toUpperCase()}/${cp.max}`, compCol: cp.col })),
      (map.internal || schema.components?.length >= 2) && { key:'_internalPct', label:'Int%' },
      map.external && { key:'_externalPct', label:'Ext%' },
      { key:'_pct',   label:'Score%' },
      { key:'_grade', label:'Grade'  },
      { key:'_gpa',   label:'GPA'    },
      { key:'_pass',  label:'Status' },
      { key:'_band',  label:'Band'   },
    ].filter(Boolean);

    let html = '<table><thead><tr>';
    cols.forEach(c => {
      const cls = APP.sortCol === c.key ? APP.sortDir : '';
      html += `<th class="${cls}" data-col="${c.key}">${c.label}</th>`;
    });
    html += '</tr></thead><tbody>';

    pageRows.forEach(row => {
      html += `<tr data-id="${row._id}">`;
      cols.forEach(c => {
        let v;
        if (c.compCol) {
          const comp = row._components.find(x => x.col === c.compCol);
          v = comp ? `${comp.value}/${comp.max}` : '—';
        } else {
          v = row[c.key] ?? '';
        }

        if (c.key === '_grade') {
          v = `<span class="g-badge g-${String(v).replace('+','p')}">${v}</span>`;
        } else if (c.key === '_pass') {
          v = row._pass ? `<span class="s-pass">Pass</span>` : `<span class="s-fail">Fail</span>`;
        } else if (c.key === '_band') {
          v = `<span style="color:${FAIL_DETECTOR.bandColor(row._band)};font-weight:600">${v}</span>`;
        } else if (typeof v === 'number') {
          v = UTILS.r(v);
        }
        html += `<td>${v}</td>`;
      });
      html += '</tr>';
    });
    html += '</tbody></table>';
    if (wrap) wrap.innerHTML = html;

    // Sort click
    wrap?.querySelectorAll('th').forEach(th => {
      th.addEventListener('click', () => {
        const col = th.dataset.col;
        APP.sortDir = APP.sortCol === col ? (APP.sortDir === 'asc' ? 'desc' : 'asc') : 'asc';
        APP.sortCol = col;
        this._renderPage();
      });
    });

    // Row click → modal
    wrap?.querySelectorAll('tr[data-id]').forEach(tr => {
      tr.addEventListener('click', () => {
        const id = parseInt(tr.dataset.id);
        const row = APP.data.find(r => r._id === id);
        if (row) this.openModal(row);
      });
    });

    this._renderPager(totalPages);
  },

  _renderPager(total) {
    const pg = UTILS.$('pager');
    if (!pg || total <= 1) { if (pg) pg.innerHTML = ''; return; }
    const cur = APP.tablePage;
    let html = '';
    if (cur > 1) html += `<button class="pg-btn" data-p="${cur-1}">‹</button>`;
    for (let i = 1; i <= total; i++) {
      if (i===1||i===total||Math.abs(i-cur)<=2)
        html += `<button class="pg-btn ${i===cur?'on':''}" data-p="${i}">${i}</button>`;
      else if (Math.abs(i-cur)===3)
        html += `<span style="padding:0 3px;color:#9ca3af">…</span>`;
    }
    if (cur < total) html += `<button class="pg-btn" data-p="${cur+1}">›</button>`;
    pg.innerHTML = html;
    pg.querySelectorAll('.pg-btn').forEach(b =>
      b.addEventListener('click', () => { APP.tablePage = +b.dataset.p; this._renderPage(); })
    );
  },

  /* ── Modal ── */
  openModal(row) {
    const schema = APP.schema;
    const name = row._name || row._roll || 'Student';
    UTILS.$('modalAv').textContent = (name||'ST').slice(0,2).toUpperCase();
    UTILS.$('modalName').textContent = name;
    UTILS.$('modalMeta').textContent = [
      row._roll    && `ID: ${row._roll}`,
      row._subject && `Subject: ${row._subject}`,
      row._dept    && `Dept: ${row._dept}`,
      row._section && `Sec: ${row._section}`,
    ].filter(Boolean).join(' • ') || '—';
    UTILS.$('modalRank').textContent = row._rank ? `#${row._rank}` : '—';

    // Dynamic marks
    const items = [];
    row._components.forEach(cp => items.push({ label: cp.col, val: cp.value, max: cp.max }));
    if (row._internalRaw !== null && !row._components.length)
      items.push({ label:'Internal', val: row._internalRaw, max: row._internalMax });
    if (row._externalRaw !== null)
      items.push({ label:'External / Sem', val: row._externalRaw, max: row._externalMax });
    if (row._totalRaw !== null)
      items.push({ label:'Total', val: row._totalRaw, max: row._totalMax });

    const colorClass = (i) => i < items.length - 1 - (row._externalRaw !== null ? 1 : 0) ? 'mf-y' : i === items.length - 1 ? 'mf-g' : 'mf-a';

    UTILS.$('modalMarks').innerHTML = items.map((it, i) => {
      const pct = it.max > 0 ? Math.min((it.val/it.max)*100, 100).toFixed(0) : 0;
      return `<div class="mr">
        <span>${it.label}</span>
        <div class="mr-track"><div class="mr-fill ${colorClass(i)}" style="width:${pct}%"></div></div>
        <span>${it.val}/${it.max} <em style="font-size:10px;opacity:.6">(${pct}%)</em></span>
      </div>`;
    }).join('');

    // Stats row
    const statsData = [
      { v: row._internalPct !== null ? row._internalPct + '%' : '—', l:'Internal%' },
      { v: row._pct + '%',  l:'Score%'  },
      { v: row._grade,      l:'Grade'   },
      { v: row._gpa,        l:'GPA'     },
      { v: row._pass ? 'Pass' : 'Fail', l:'Status', c: row._pass ? '#16a34a' : '#dc2626' },
    ];
    UTILS.$('modalStatsRow').innerHTML = statsData.map(s =>
      `<div class="mstat"><span class="mstat-v" ${s.c?`style="color:${s.c}"`:''}>${s.v}</span><span class="mstat-l">${s.l}</span></div>`
    ).join('');

    // Radar
    if (APP.modalChart) { APP.modalChart.destroy(); APP.modalChart = null; }
    const radarLabels = [], radarData = [];
    if (row._components.length) {
      row._components.forEach(c => { radarLabels.push(c.col); radarData.push(c.pct); });
    } else {
      if (row._internalPct !== null) { radarLabels.push('Internal'); radarData.push(row._internalPct); }
      if (row._externalPct !== null) { radarLabels.push('External'); radarData.push(row._externalPct); }
    }
    if (!radarLabels.length) { radarLabels.push('Score'); radarData.push(row._pct); }

    const rc = UTILS.$('modalRadar');
    if (rc) {
      APP.modalChart = new Chart(rc.getContext('2d'), {
        type:'radar',
        data:{ labels:radarLabels, datasets:[{ label:name, data:radarData,
          backgroundColor:'rgba(250,204,21,.15)', borderColor:'rgba(250,204,21,.9)',
          pointBackgroundColor:'rgba(234,179,8,.9)', borderWidth:2 }] },
        options:{ responsive:true, maintainAspectRatio:false,
          scales:{ r:{ beginAtZero:true, max:100,
            ticks:{ display:false, backdropColor:'transparent' },
            grid:{ color:'rgba(0,0,0,.06)' },
            pointLabels:{ color:'#6b7280', font:{ family:'Inter', size:11 } } } },
          plugins:{ legend:{ display:false },
            tooltip:{ backgroundColor:'rgba(255,255,255,.97)', borderColor:'rgba(250,204,21,.3)',
              borderWidth:1, titleColor:'#111827', bodyColor:'#6b7280',
              bodyFont:{ family:'JetBrains Mono', size:11 } } } },
      });
    }

    // Strengths / weaknesses
    const scored = radarLabels.map((l, i) => ({ name:l, pct:radarData[i] }));
    UTILS.$('modalStr').textContent =
      scored.filter(s=>s.pct>=70).map(s=>`${s.name} (${s.pct.toFixed(0)}%)`).join(', ') || 'No standout strengths';
    UTILS.$('modalWk').textContent =
      scored.filter(s=>s.pct<50).map(s=>`${s.name} (${s.pct.toFixed(0)}%)`).join(', ') || 'No major weaknesses';

    UTILS.show('modalBg');
    document.body.style.overflow = 'hidden';
  },

  closeModal() { UTILS.hide('modalBg'); document.body.style.overflow = ''; },
};

/* ================================================================
   EXPORT
   ================================================================ */
const EXPORT = {
  csv() {
    if (!APP.data.length) { UTILS.toast('No data to export'); return; }
    const schema = APP.schema; const map = schema.map || {};
    const baseCols = ['_name','_roll','_subject','_dept','_section'];
    const baseLabels = ['Name','Roll/ID','Subject','Dept','Section'];
    const compCols = schema.components.map(cp => ({ col:`__comp_${cp.col}`, label:`${cp.role}/${cp.max}`, cpCol:cp.col }));
    const calcCols = [
      { col:'_internalPct', label:'Internal%' },
      { col:'_externalPct', label:'External%' },
      { col:'_pct',         label:'Score%'    },
      { col:'_grade',       label:'Grade'     },
      { col:'_gpa',         label:'GPA'       },
      { col:'_pass',        label:'Status'    },
      { col:'_band',        label:'Band'      },
      { col:'_rank',        label:'Rank'      },
    ];

    const headers = [...baseLabels, ...compCols.map(c=>c.label), ...calcCols.map(c=>c.label)];
    const rows = APP.filtered.map(r => [
      ...baseCols.map(c => r[c] ?? ''),
      ...compCols.map(c => { const cp = r._components.find(x => x.col === c.cpCol); return cp ? cp.value : ''; }),
      ...calcCols.map(c => c.col === '_pass' ? (r._pass ? 'Pass' : 'Fail') : (r[c.col] ?? '')),
    ]);

    const csv = [headers, ...rows].map(row =>
      row.map(v => `"${String(v??'').replace(/"/g,'""')}"`).join(',')
    ).join('\n');
    this._dl('academic_report.csv','text/csv', csv);
    UTILS.toast('CSV exported!');
  },

  excel() {
    if (!APP.data.length) { UTILS.toast('No data'); return; }
    const schema = APP.schema;
    const data = APP.filtered.map(r => {
      const row = {};
      if (r._name) row.Name = r._name;
      if (r._roll) row['Roll/ID'] = r._roll;
      if (r._subject) row.Subject = r._subject;
      if (r._dept)    row.Department = r._dept;
      if (r._section) row.Section = r._section;
      schema.components?.forEach(cp => {
        const c = r._components.find(x => x.col === cp.col);
        row[`${cp.col}/${cp.max}`] = c ? c.value : '';
      });
      if (r._internalPct !== null) row['Internal%'] = r._internalPct;
      if (r._externalPct !== null) row['External%'] = r._externalPct;
      row['Score%'] = r._pct;
      row.Grade = r._grade;
      row.GPA   = r._gpa;
      row.Status = r._pass ? 'Pass' : 'Fail';
      row.Band  = r._band;
      row.Rank  = r._rank;
      return row;
    });
    const S = APP.stats;
    const statsArr = [
      ['Metric','Value'],
      ['Total',S.total],['Pass',S.passCount],['Fail',S.failCount],
      ['Pass Rate',S.passRate+'%'],['Average',S.avg+'%'],['Median',S.median+'%'],
      ['Std Dev',S.stdDev],['High',S.high+'%'],['Low',S.low+'%'],
      ['Avg GPA',S.gpaStats?.avg],['Schema Mode',schema.mode],
      ...Object.entries(S.bandCounts).map(([k,v])=>[k,v]),
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), 'Student Data');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(statsArr), 'Statistics');
    XLSX.writeFile(wb, 'academic_report.xlsx');
    UTILS.toast('Excel exported!');
  },

  charts() {
    const ids = ['cSubjectAvg','cSubjectPF','cPassFail','cGrade','cScoreBands',
      'cGPA','cBandPie','cIntVsExt','cComponents','cDept','cTrend'];
    let n = 0;
    ids.forEach(id => {
      const c = UTILS.$(id); if (!c) return;
      const a = document.createElement('a');
      a.download = `chart_${id}.png`;
      a.href = c.toDataURL('image/png');
      a.click(); n++;
    });
    UTILS.toast(`${n} charts exported!`);
  },

  pdf() { UTILS.toast('Opening print dialog…'); setTimeout(() => window.print(), 500); },

  _dl(name, type, content) {
    const blob = new Blob([content], { type });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name; a.click();
    URL.revokeObjectURL(a.href);
  },
};

/* ================================================================
   BOOTSTRAP — wire up all events
   ================================================================ */
(function boot() {
  const $ = UTILS.$;

  /* Navbar scroll */
  window.addEventListener('scroll', () => {
    const nb = $('navbar');
    if (nb) nb.style.boxShadow = window.scrollY > 60 ? '0 2px 20px rgba(0,0,0,.08)' : 'none';
  });

  /* Hamburger */
  $('hamburger')?.addEventListener('click', () => $('navLinks')?.classList.toggle('hidden'));

  /* Upload wiring */
  $('navUploadBtn')?.addEventListener('click', () => $('fileInput')?.click());
  $('heroUploadBtn')?.addEventListener('click', () => $('fileInput')?.click());
  $('uploadTrigger')?.addEventListener('click', () => $('fileInput')?.click());
  $('changeFileBtn')?.addEventListener('click', () => {
    $('fileInput').value = '';
    $('uploadDone')?.classList.add('hidden');
    $('uploadIdle')?.classList.remove('hidden');
  });

  /* Drag & drop */
  const zone = $('uploadCard');
  zone?.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone?.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone?.addEventListener('drop', e => {
    e.preventDefault(); zone.classList.remove('drag-over');
    if (e.dataTransfer.files[0]) processFile(e.dataTransfer.files[0]);
  });

  $('fileInput')?.addEventListener('change', e => {
    if (e.target.files[0]) processFile(e.target.files[0]);
  });

  /* Table search */
  $('tblSearch')?.addEventListener('input', UTILS.debounce(() => {
    APP.tblSearchQ = ($('tblSearch')?.value ?? '').toLowerCase();
    APP.tablePage = 1;
    UI._renderPage();
  }, 260));

  $('tblPageSize')?.addEventListener('change', () => {
    APP.tablePageSize = +($('tblPageSize')?.value || 10);
    APP.tablePage = 1;
    UI._renderPage();
  });

  /* Filters */
  ['fSubject','fDept','fSection','fGrade','fStatus'].forEach(id =>
    $(id)?.addEventListener('change', () => APP.applyFilters())
  );
  $('resetFilters')?.addEventListener('click', () => {
    ['fSubject','fDept','fSection','fGrade','fStatus'].forEach(id => { if ($(id)) $(id).value = ''; });
    APP.applyFilters();
  });

  /* Chart type toggles */
  document.querySelectorAll('[data-ckey="subjectAvg"]').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('[data-ckey="subjectAvg"]').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      CHART_ENGINE.subjectAvg(APP.stats, btn.dataset.type);
    });
  });

  /* Modal close */
  $('modalX')?.addEventListener('click', () => UI.closeModal());
  $('modalBg')?.addEventListener('click', e => { if (e.target === $('modalBg')) UI.closeModal(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape') UI.closeModal(); });
})();

/* ── File processor ── */
function processFile(file) {
  if (!file.name.match(/\.(xlsx|xls)$/i)) { UTILS.toast('Please upload .xlsx or .xls'); return; }
  UTILS.loader('Reading Excel file…');
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      const rows = PARSER.parse(ev.target.result);
      if (!rows.length) { UTILS.loaderOff(); UTILS.toast('File appears empty'); return; }
      APP.run(rows, file.name, file.size);
    } catch(e) {
      UTILS.loaderOff(); UTILS.toast('Could not parse file'); console.error(e);
    }
  };
  reader.readAsArrayBuffer(file);
}

/* ── Global exports (for HTML onclick attributes) ── */
window.EXPORT = EXPORT;
