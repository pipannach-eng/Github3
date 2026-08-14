(function () {
  const dashboardData = window.DASHBOARD_DATA;
  if (!dashboardData) return;

  const PASS_LABEL = "ผ่าน";
  const NO_ISSUE_LABEL = "ไม่มีประเด็น";
  const NO_ISSUE_DETAIL = "ไม่พบจำนวนข้อที่ต้องปรับปรุงในระเบียนนี้";
  const NO_DATA_MESSAGE = "ไม่พบข้อมูลในตัวกรองปัจจุบัน";
  const NO_YEAR_DATA_MESSAGE = "ไม่พบข้อมูลรายปีในตัวกรองปัจจุบัน";
  const PASS_TARGET_RATE = 95;

  const records = dashboardData.records.map((record) => ({
    ...record,
    assessment_year_be: Number(record.assessment_year_be),
    assessment_year_ce: Number(record.assessment_year_ce),
    ImproveScore: Number(record.ImproveScore),
    center_name: record.center_name || record.center_code_anon,
    center_alias: record.center_code_anon,
  }));

  const indicatorReference = dashboardData.indicator_reference;
  const indicatorByColumn = Object.fromEntries(
    indicatorReference.map((item) => [item.feature_column, item])
  );
  const indicatorColumns = indicatorReference.map((item) => item.feature_column);
  const years = [...new Set(records.map((record) => record.assessment_year_be))].sort((a, b) => a - b);
  const qualityLevels = ["A", "B", "C", "D"];
  const schoolOptions = Object.values(
    records.reduce((acc, record) => {
      if (!acc[record.center_code_anon]) {
        acc[record.center_code_anon] = {
          center_code_anon: record.center_code_anon,
          center_name: record.center_name,
          center_alias: record.center_code_anon,
        };
      }
      return acc;
    }, {})
  ).sort((a, b) => a.center_alias.localeCompare(b.center_alias, "th"));

  const state = {
    selectedYears: new Set(years),
    selectedQuality: new Set(qualityLevels),
    selectedCenterCode: schoolOptions[0]?.center_code_anon || "",
  };

  const elements = {
    sourceSheet: document.getElementById("source-sheet"),
    totalRecords: document.getElementById("total-records"),
    generatedAt: document.getElementById("generated-at"),
    headerPassRate: document.getElementById("header-pass-rate"),
    headerFailRate: document.getElementById("header-fail-rate"),
    headerQualityD: document.getElementById("header-quality-d"),
    headerAssessedRecords: document.getElementById("header-assessed-records"),
    yearFilters: document.getElementById("year-filters"),
    qualityFilters: document.getElementById("quality-filters"),
    resetYears: document.getElementById("reset-years"),
    resetQuality: document.getElementById("reset-quality"),
    schoolSearch: document.getElementById("school-search"),
    schoolOptions: document.getElementById("school-options"),
    selectedSchoolName: document.getElementById("selected-school-name"),
    selectedSchoolCode: document.getElementById("selected-school-code"),
    kpiTotal: document.getElementById("kpi-total"),
    kpiPassRate: document.getElementById("kpi-pass-rate"),
    kpiQualityD: document.getElementById("kpi-quality-d"),
    execPassPercent: document.getElementById("exec-pass-percent"),
    execPassCount: document.getElementById("exec-pass-count"),
    execWatchPercent: document.getElementById("exec-watch-percent"),
    execWatchCount: document.getElementById("exec-watch-count"),
    execCriticalPercent: document.getElementById("exec-critical-percent"),
    execCriticalCount: document.getElementById("exec-critical-count"),
    execNoDataPercent: document.getElementById("exec-nodata-percent"),
    execNoDataCount: document.getElementById("exec-nodata-count"),
    execStatusDonut: document.getElementById("exec-status-donut"),
    execTrendChart: document.getElementById("exec-trend-chart"),
    execCategoryChart: document.getElementById("exec-category-chart"),
    execRankingPanel: document.getElementById("exec-ranking-panel"),
    execCriticalIndicators: document.getElementById("exec-critical-indicators"),
    execCompositionChart: document.getElementById("exec-composition-chart"),
    execGeoMode: document.getElementById("exec-geo-mode"),
    execGeoAnalysis: document.getElementById("exec-geo-analysis"),
    execAlerts: document.getElementById("exec-alerts"),
    execValidationList: document.getElementById("exec-validation-list"),
    policyRiskCenters: document.getElementById("policy-risk-centers"),
    policyWatchCenters: document.getElementById("policy-watch-centers"),
    policyAvgScore: document.getElementById("policy-avg-score"),
    policyResourceTheme: document.getElementById("policy-resource-theme"),
    policyPriorityList: document.getElementById("policy-priority-list"),
    policyIssueList: document.getElementById("policy-issue-list"),
    policyActionList: document.getElementById("policy-action-list"),
    rankingChart: document.getElementById("ranking-chart"),
    yearlyChart: document.getElementById("yearly-chart"),
    schoolTimelineChart: document.getElementById("school-timeline-chart"),
    schoolTimelineSummary: document.getElementById("school-timeline-summary"),
    topIndicators: document.getElementById("top-indicators"),
    trackingTableBody: document.getElementById("tracking-table-body"),
  };

  function formatNumber(value) {
    return new Intl.NumberFormat("th-TH").format(value);
  }

  function formatPercent(value) {
    return `${value.toFixed(1)}%`;
  }

  function truncate(text, max = 88) {
    if (!text) return "-";
    return text.length > max ? `${text.slice(0, max - 1)}...` : text;
  }

  function getCssVar(name, fallback) {
    const value = getComputedStyle(document.documentElement).getPropertyValue(name).trim();
    return value || fallback;
  }

  function emptyState(message) {
    return `<div class="empty-state">${message}</div>`;
  }

  function getFilteredRecords() {
    return records.filter(
      (record) =>
        state.selectedYears.has(record.assessment_year_be) &&
        state.selectedQuality.has(record.Quality_level)
    );
  }

  function getSchoolDisplay(option) {
    return option.center_alias;
  }

  function findSchoolOptionByInput(inputValue) {
    const normalized = inputValue.trim();
    const exactMatch = schoolOptions.find(
      (option) =>
        option.center_alias === normalized ||
        option.center_code_anon === normalized ||
        getSchoolDisplay(option) === normalized
    );
    if (exactMatch) return exactMatch;

    const normalizedLower = normalized.toLocaleLowerCase("th");
    return schoolOptions.find(
      (option) =>
        option.center_alias.toLocaleLowerCase("th").includes(normalizedLower) ||
        option.center_code_anon.toLocaleLowerCase("th").includes(normalizedLower)
    );
  }

  function createChip(label, active, onClick, extraClass = "") {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `chip${extraClass ? ` ${extraClass}` : ""}${active ? " active" : ""}`;
    button.textContent = label;
    button.addEventListener("click", onClick);
    return button;
  }

  function renderFilters() {
    elements.yearFilters.innerHTML = "";
    years.forEach((year) => {
      elements.yearFilters.appendChild(
        createChip(String(year), state.selectedYears.has(year), () => {
          if (state.selectedYears.has(year)) {
            state.selectedYears.delete(year);
          } else {
            state.selectedYears.add(year);
          }
          if (state.selectedYears.size === 0) {
            years.forEach((item) => state.selectedYears.add(item));
          }
          render();
        })
      );
    });

    elements.qualityFilters.innerHTML = "";
    qualityLevels.forEach((level) => {
      elements.qualityFilters.appendChild(
        createChip(level, state.selectedQuality.has(level), () => {
          if (state.selectedQuality.has(level)) {
            state.selectedQuality.delete(level);
          } else {
            state.selectedQuality.add(level);
          }
          if (state.selectedQuality.size === 0) {
            qualityLevels.forEach((item) => state.selectedQuality.add(item));
          }
          render();
        }, `level-${level.toLowerCase()}`)
      );
    });
  }

  function renderSchoolSearchOptions() {
    elements.schoolOptions.innerHTML = "";
    schoolOptions.forEach((option) => {
      const item = document.createElement("option");
      item.value = getSchoolDisplay(option);
      elements.schoolOptions.appendChild(item);
    });
  }

  function computeIndicatorFrequency(filteredRecords) {
    const frequency = Object.fromEntries(indicatorColumns.map((column) => [column, 0]));
    filteredRecords.forEach((record) => {
      indicatorColumns.forEach((column) => {
        frequency[column] += Number(record[column] || 0);
      });
    });
    return frequency;
  }

  function getTopIndicatorForRecord(record, frequencyMap) {
    const matched = indicatorColumns
      .filter((column) => Number(record[column]) === 1)
      .map((column) => ({
        column,
        count: frequencyMap[column],
        item: indicatorByColumn[column],
      }))
      .sort((a, b) => b.count - a.count || a.item.item_code.localeCompare(b.item.item_code, "th"));

    if (matched.length === 0) {
      return { label: NO_ISSUE_LABEL, detail: NO_ISSUE_DETAIL };
    }

    return {
      label: matched[0].item.item_code,
      detail: matched[0].item.indicator_text,
    };
  }

  function renderKpis(filteredRecords) {
    const total = filteredRecords.length;
    const passCount = filteredRecords.filter((item) => item.target_label === PASS_LABEL).length;
    const qualityD = filteredRecords.filter((item) => item.Quality_level === "D").length;
    const passRate = total === 0 ? 0 : (passCount / total) * 100;

    elements.kpiTotal.textContent = formatNumber(total);
    elements.kpiPassRate.textContent = formatPercent(passRate);
    elements.kpiQualityD.textContent = formatNumber(qualityD);
  }

  function getPolicyGroupField(filteredRecords) {
    const candidates = [
      { key: "affiliation_3", label: "เขต/พื้นที่" },
      { key: "affiliation_2", label: "สังกัด/กลุ่มหน่วยงาน" },
      { key: "affiliation_1", label: "สังกัดหลัก" },
      { key: "ministry", label: "กระทรวง/หน่วยงาน" },
    ];
    return candidates.find(({ key }) =>
      filteredRecords.some((record) => record[key] && String(record[key]).trim() !== "")
    ) || { key: "center_code_anon", label: "สถานศึกษาเป้าหมาย" };
  }

  function getActionTheme(item) {
    const code = item?.item_code || "";
    const text = item?.indicator_text || "";
    if (code.startsWith("1.2")) {
      return {
        title: "พัฒนาศักยภาพบุคลากร",
        detail: "จัดอบรม/พี่เลี้ยงวิชาการ และวางแผนสนับสนุนครูหรือผู้ดูแลเด็กในกลุ่มเป้าหมาย",
      };
    }
    if (code.startsWith("1.3")) {
      return {
        title: "ยกระดับความปลอดภัยและสภาพแวดล้อม",
        detail: "เร่งตรวจความพร้อมอาคาร พื้นที่เล่น แผนฉุกเฉิน และอุปกรณ์ป้องกันตามความเสี่ยงของพื้นที่",
      };
    }
    if (code.startsWith("1.4") || code === "3.1.3ข" || text.includes("สุข")) {
      return {
        title: "บูรณาการสุขภาพเด็กเชิงรุก",
        detail: "เชื่อมหน่วยบริการสุขภาพในพื้นที่ ตรวจคัดกรอง ติดตาม และส่งเสริมพฤติกรรมสุขภาพอย่างต่อเนื่อง",
      };
    }
    if (code.startsWith("2.")) {
      return {
        title: "เสริมคุณภาพกระบวนการดูแลและจัดประสบการณ์",
        detail: "สนับสนุน coaching ครู แผนจัดประสบการณ์ และระบบติดตามคุณภาพการดูแลเด็กปฐมวัย",
      };
    }
    if (code.startsWith("3.5") || code.startsWith("3.6")) {
      return {
        title: "สนับสนุนการเรียนรู้ผ่านการเล่น",
        detail: "จัดสื่อการเรียนรู้เชิงรูปธรรม กิจกรรมภาษา/คณิตศาสตร์ และแนวทาง play-based learning",
      };
    }
    return {
      title: "จัดทีมช่วยเหลือตามมาตรฐานที่พบ",
      detail: "ใช้รายการมาตรฐานที่พบมากเพื่อกำหนดหน่วยสนับสนุน งบประมาณ และแผนติดตามเฉพาะจุด",
    };
  }

  function renderPolicyAction(filteredRecords, frequencyMap) {
    if (!elements.policyPriorityList) return;

    const total = filteredRecords.length;
    const riskRecords = filteredRecords.filter(
      (record) => record.Quality_level === "D" || record.ImproveScore >= 16
    );
    const watchRecords = filteredRecords.filter(
      (record) => record.Quality_level === "C" || (record.ImproveScore >= 8 && record.ImproveScore <= 15)
    );
    const avgScore = total === 0
      ? 0
      : filteredRecords.reduce((sum, record) => sum + record.ImproveScore, 0) / total;

    const topIssues = indicatorReference
      .map((item) => ({ ...item, count: frequencyMap[item.feature_column] || 0 }))
      .filter((item) => item.count > 0)
      .sort((a, b) => b.count - a.count || a.item_code.localeCompare(b.item_code, "th"));

    const topTheme = topIssues.length ? getActionTheme(topIssues[0]).title : "รักษาคุณภาพและติดตามต่อเนื่อง";

    elements.policyRiskCenters.textContent = formatNumber(riskRecords.length);
    elements.policyWatchCenters.textContent = formatNumber(watchRecords.length);
    elements.policyAvgScore.textContent = avgScore.toFixed(1);
    elements.policyResourceTheme.textContent = topTheme;

    if (total === 0) {
      elements.policyPriorityList.innerHTML = emptyState(NO_DATA_MESSAGE);
      elements.policyIssueList.innerHTML = emptyState(NO_DATA_MESSAGE);
      elements.policyActionList.innerHTML = emptyState(NO_DATA_MESSAGE);
      return;
    }

    const groupInfo = getPolicyGroupField(filteredRecords);
    const grouped = Object.values(
      filteredRecords.reduce((acc, record) => {
        const key = String(record[groupInfo.key] || record.center_code_anon || "ไม่ระบุ").trim() || "ไม่ระบุ";
        if (!acc[key]) {
          acc[key] = { key, total: 0, risk: 0, watch: 0, improveSum: 0, maxScore: 0 };
        }
        acc[key].total += 1;
        acc[key].improveSum += record.ImproveScore;
        acc[key].maxScore = Math.max(acc[key].maxScore, record.ImproveScore);
        if (record.Quality_level === "D" || record.ImproveScore >= 16) acc[key].risk += 1;
        if (record.Quality_level === "C" || (record.ImproveScore >= 8 && record.ImproveScore <= 15)) acc[key].watch += 1;
        return acc;
      }, {})
    )
      .map((item) => ({
        ...item,
        avgScore: item.total === 0 ? 0 : item.improveSum / item.total,
        priorityScore: item.risk * 6 + item.watch * 3 + item.improveSum / Math.max(item.total, 1),
      }))
      .sort((a, b) => b.priorityScore - a.priorityScore || b.avgScore - a.avgScore)
      .slice(0, 5);

    elements.policyPriorityList.innerHTML = grouped
      .map((item, index) => `
        <div class="policy-item">
          <span class="policy-rank">${index + 1}</span>
          <div>
            <b>${item.key}</b>
            <span>${groupInfo.label}: รวม ${formatNumber(item.total)} ระเบียน | ควรเร่งดูแล ${formatNumber(item.risk)} | เฝ้าระวัง ${formatNumber(item.watch)}</span>
          </div>
          <span class="policy-score">เฉลี่ย ${item.avgScore.toFixed(1)}</span>
        </div>
      `)
      .join("");

    elements.policyIssueList.innerHTML = topIssues.length
      ? topIssues.slice(0, 5).map((item, index) => `
        <div class="policy-item">
          <span class="policy-rank">${index + 1}</span>
          <div>
            <b>${item.item_code}</b>
            <span>${truncate(item.indicator_text, 108)}</span>
          </div>
          <span class="policy-score">${formatNumber(item.count)} ครั้ง</span>
        </div>
      `).join("")
      : emptyState("ยังไม่พบมาตรฐานที่ต้องปรับปรุงในตัวกรองนี้");

    const actionMap = new Map();
    topIssues.forEach((item) => {
      const action = getActionTheme(item);
      const current = actionMap.get(action.title) || { ...action, count: 0, examples: [] };
      current.count += item.count;
      if (current.examples.length < 2) current.examples.push(item.item_code);
      actionMap.set(action.title, current);
    });

    const actions = [...actionMap.values()].sort((a, b) => b.count - a.count).slice(0, 4);
    elements.policyActionList.innerHTML = actions.length
      ? actions.map((action) => `
        <div class="policy-action-item">
          <b>${action.title}</b>
          <span>${action.detail} <strong>มาตรฐานอ้างอิง:</strong> ${action.examples.join(", ")} | พบ ${formatNumber(action.count)} ครั้ง</span>
        </div>
      `).join("")
      : `
        <div class="policy-action-item">
          <b>รักษาคุณภาพและติดตามต่อเนื่อง</b>
          <span>ไม่พบประเด็นต้องปรับปรุงในตัวกรองนี้ ควรใช้เป็นกลุ่มต้นแบบหรือพื้นที่แลกเปลี่ยนเรียนรู้</span>
        </div>
      `;
  }

  function polarToCartesian(cx, cy, radius, angleInDegrees) {
    const angleInRadians = ((angleInDegrees - 90) * Math.PI) / 180;
    return {
      x: cx + radius * Math.cos(angleInRadians),
      y: cy + radius * Math.sin(angleInRadians),
    };
  }

  function describeArc(cx, cy, radius, startAngle, endAngle) {
    const start = polarToCartesian(cx, cy, radius, endAngle);
    const end = polarToCartesian(cx, cy, radius, startAngle);
    const largeArcFlag = endAngle - startAngle <= 180 ? "0" : "1";
    return `M ${start.x} ${start.y} A ${radius} ${radius} 0 ${largeArcFlag} 0 ${end.x} ${end.y}`;
  }

  function getStatusSummary(filteredRecords) {
    const total = filteredRecords.length;
    const noData = filteredRecords.filter((record) =>
      ["missing", "invalid", "unknown_marker_17"].includes(String(record.improvement_data_status || ""))
    ).length;
    const critical = filteredRecords.filter((record) => record.Quality_level === "D").length;
    const watch = filteredRecords.filter((record) => record.Quality_level === "C").length;
    const pass = Math.max(total - noData - critical - watch, 0);
    return { total, pass, watch, critical, noData };
  }

  function getTrendSymbol(current, previous, lowerIsBetter = false) {
    if (previous == null || Math.abs(current - previous) < 0.05) {
      return { symbol: "■", className: "trend-flat", label: "ทรงตัว" };
    }
    const improved = lowerIsBetter ? current < previous : current > previous;
    return improved
      ? { symbol: "▲", className: "trend-up", label: "ดีขึ้น" }
      : { symbol: "▼", className: "trend-down", label: "แย่ลง" };
  }

  function getYearRows(sourceRecords) {
    return years.map((year) => {
      const items = sourceRecords.filter((record) => record.assessment_year_be === year);
      const total = items.length;
      const passCount = items.filter((record) => record.target_label === PASS_LABEL).length;
      const dCount = items.filter((record) => record.Quality_level === "D").length;
      const avgScore = total === 0 ? 0 : items.reduce((sum, record) => sum + record.ImproveScore, 0) / total;
      return {
        year,
        total,
        passRate: total === 0 ? 0 : (passCount / total) * 100,
        dRate: total === 0 ? 0 : (dCount / total) * 100,
        avgScore,
      };
    }).filter((row) => row.total > 0);
  }

  function renderExecutiveStatus(filteredRecords) {
    const summary = getStatusSummary(filteredRecords);
    const total = summary.total || 1;
    const rows = [
      { key: "pass", label: "ผ่านเกณฑ์", value: summary.pass, color: "#2fa66a" },
      { key: "watch", label: "เฝ้าระวัง", value: summary.watch, color: "#f2a93b" },
      { key: "critical", label: "ต้องปรับปรุง", value: summary.critical, color: "#d94e4e" },
      { key: "noData", label: "ไม่มีข้อมูล", value: summary.noData, color: "#a9b5c2" },
    ];

    elements.execPassPercent.textContent = formatPercent((summary.pass / total) * 100);
    elements.execPassCount.textContent = `${formatNumber(summary.pass)} ระเบียน`;
    elements.execWatchPercent.textContent = formatPercent((summary.watch / total) * 100);
    elements.execWatchCount.textContent = `${formatNumber(summary.watch)} ระเบียน`;
    elements.execCriticalPercent.textContent = formatPercent((summary.critical / total) * 100);
    elements.execCriticalCount.textContent = `${formatNumber(summary.critical)} ระเบียน`;
    elements.execNoDataPercent.textContent = formatPercent((summary.noData / total) * 100);
    elements.execNoDataCount.textContent = `${formatNumber(summary.noData)} ระเบียน`;

    if (!summary.total) {
      elements.execStatusDonut.innerHTML = emptyState(NO_DATA_MESSAGE);
      return;
    }

    let angle = 0;
    const arcs = rows.map((row) => {
      const sweep = (row.value / summary.total) * 360;
      const path = sweep <= 0
        ? ""
        : `<path d="${describeArc(96, 96, 66, angle, angle + sweep)}" fill="none" stroke="${row.color}" stroke-width="24" stroke-linecap="round"></path>`;
      angle += sweep;
      return path;
    }).join("");

    const legend = rows.map((row) => `
      <div class="exec-legend-row">
        <span class="exec-dot" style="background:${row.color}"></span>
        <div><b>${row.label}</b><span>${formatPercent((row.value / summary.total) * 100)}</span></div>
        <strong>${formatNumber(row.value)}</strong>
      </div>
    `).join("");

    elements.execStatusDonut.innerHTML = `
      <svg class="exec-svg" viewBox="0 0 192 192" role="img" aria-label="สถานะภาพรวม">
        <circle cx="96" cy="96" r="66" fill="none" stroke="#edf3f8" stroke-width="24"></circle>
        ${arcs}
        <text x="96" y="90" text-anchor="middle" font-size="24" font-weight="800" fill="#18365f">${formatPercent((summary.pass / summary.total) * 100)}</text>
        <text x="96" y="114" text-anchor="middle" font-size="12" fill="#66788f">ผ่านเกณฑ์</text>
      </svg>
      <div class="exec-legend">${legend}</div>
    `;
  }

  function renderExecutiveTrend(filteredRecords) {
    const rows = getYearRows(filteredRecords).slice(-3);
    if (!rows.length) {
      elements.execTrendChart.innerHTML = emptyState(NO_YEAR_DATA_MESSAGE);
      return;
    }

    const width = 430;
    const height = 190;
    const left = 42;
    const top = 22;
    const chartHeight = 118;
    const plotWidth = width - left - 18;
    const step = rows.length === 1 ? 0 : plotWidth / (rows.length - 1);
    const points = rows.map((row, index) => ({
      x: left + index * step,
      y: top + chartHeight - (row.passRate / 100) * chartHeight,
      row,
    }));
    const dPoints = rows.map((row, index) => ({
      x: left + index * step,
      y: top + chartHeight - (row.dRate / 100) * chartHeight,
      row,
    }));
    const passPath = points.map((point, index) => `${index ? "L" : "M"} ${point.x} ${point.y}`).join(" ");
    const dPath = dPoints.map((point, index) => `${index ? "L" : "M"} ${point.x} ${point.y}`).join(" ");
    const grid = [0, 50, 100].map((value) => {
      const y = top + chartHeight - (value / 100) * chartHeight;
      return `<line x1="${left}" x2="${width - 12}" y1="${y}" y2="${y}" stroke="#edf3f8"></line><text x="${left - 8}" y="${y + 4}" text-anchor="end" font-size="11" fill="#66788f">${value}%</text>`;
    }).join("");
    const marks = points.map((point, index) => `
      <circle cx="${point.x}" cy="${point.y}" r="5" fill="#4371ad" stroke="#fff" stroke-width="2"></circle>
      <circle cx="${dPoints[index].x}" cy="${dPoints[index].y}" r="5" fill="#d94e4e" stroke="#fff" stroke-width="2"></circle>
      <text x="${point.x}" y="${height - 18}" text-anchor="middle" font-size="11" fill="#23364e" font-weight="800">${point.row.year}</text>
    `).join("");

    elements.execTrendChart.innerHTML = `
      <svg class="exec-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="แนวโน้ม 3 ปีล่าสุด">
        ${grid}
        <path d="${passPath}" fill="none" stroke="#4371ad" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"></path>
        <path d="${dPath}" fill="none" stroke="#d94e4e" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"></path>
        ${marks}
        <rect x="${width - 164}" y="6" width="10" height="10" rx="2" fill="#4371ad"></rect>
        <text x="${width - 149}" y="15" font-size="11" fill="#23364e">อัตราผ่าน</text>
        <rect x="${width - 82}" y="6" width="10" height="10" rx="2" fill="#d94e4e"></rect>
        <text x="${width - 67}" y="15" font-size="11" fill="#23364e">ระดับ D</text>
      </svg>
    `;
  }

  function renderExecutiveCategory(filteredRecords) {
    if (!filteredRecords.length) {
      elements.execCategoryChart.innerHTML = emptyState(NO_DATA_MESSAGE);
      return;
    }

    const groupInfo = getPolicyGroupField(filteredRecords);
    const grouped = Object.values(filteredRecords.reduce((acc, record) => {
      const key = String(record[groupInfo.key] || record.center_code_anon || "ไม่ระบุ").trim() || "ไม่ระบุ";
      if (!acc[key]) acc[key] = { key, total: 0, pass: 0, d: 0 };
      acc[key].total += 1;
      if (record.target_label === PASS_LABEL) acc[key].pass += 1;
      if (record.Quality_level === "D") acc[key].d += 1;
      return acc;
    }, {}))
      .map((item) => ({
        ...item,
        passRate: item.total === 0 ? 0 : (item.pass / item.total) * 100,
        variance: item.total === 0 ? 0 : (item.pass / item.total) * 100 - PASS_TARGET_RATE,
      }))
      .sort((a, b) => a.passRate - b.passRate || b.total - a.total)
      .slice(0, 8);

    elements.execCategoryChart.innerHTML = grouped.map((item) => `
      <div class="exec-bar-row">
        <div class="exec-bar-label" title="${item.key}">${item.key}</div>
        <div class="exec-bar-track">
          <div class="exec-bar-fill" style="width:${Math.max(1, item.passRate)}%; background:${item.passRate < PASS_TARGET_RATE ? "linear-gradient(90deg,#d94e4e,#f2a93b)" : "linear-gradient(90deg,#2f7bd3,#72b8f7)"}"></div>
          <span class="exec-target-marker" style="left:${PASS_TARGET_RATE}%"></span>
        </div>
        <div class="exec-bar-value">${formatPercent(item.passRate)}</div>
      </div>
    `).join("");
  }

  function getCenterTrend(centerRecords) {
    const sorted = centerRecords.slice().sort((a, b) => a.assessment_year_be - b.assessment_year_be);
    if (sorted.length < 2) return getTrendSymbol(0, 0, true);
    const current = sorted[sorted.length - 1].ImproveScore;
    const previous = sorted[sorted.length - 2].ImproveScore;
    return getTrendSymbol(current, previous, true);
  }

  function renderExecutiveRankings(filteredRecords) {
    if (!filteredRecords.length) {
      elements.execRankingPanel.innerHTML = emptyState(NO_DATA_MESSAGE);
      return;
    }

    const latestByCenter = Object.values(records.reduce((acc, record) => {
      if (!state.selectedYears.has(record.assessment_year_be)) return acc;
      if (!acc[record.center_code_anon] || record.assessment_year_be > acc[record.center_code_anon].assessment_year_be) {
        acc[record.center_code_anon] = record;
      }
      return acc;
    }, {})).filter((record) => state.selectedQuality.has(record.Quality_level));
    const centerGroups = records.reduce((acc, record) => {
      if (!acc[record.center_code_anon]) acc[record.center_code_anon] = [];
      acc[record.center_code_anon].push(record);
      return acc;
    }, {});
    const ranked = latestByCenter.map((record) => ({
      ...record,
      trend: getCenterTrend(centerGroups[record.center_code_anon] || [record]),
    }));
    const best = ranked.slice().sort((a, b) => a.ImproveScore - b.ImproveScore || a.center_code_anon.localeCompare(b.center_code_anon)).slice(0, 5);
    const bottom = ranked.slice().sort((a, b) => b.ImproveScore - a.ImproveScore || a.center_code_anon.localeCompare(b.center_code_anon)).slice(0, 5);

    const renderRows = (items, attention = false) => items.map((item, index) => `
      <div class="exec-rank-row${attention && item.ImproveScore > 0 ? " attention" : ""}">
        <span class="exec-rank-no">${index + 1}</span>
        <div><b>${item.center_code_anon}</b><span>${formatNumber(item.ImproveScore)} จำนวนข้อที่ต้องปรับปรุง | ระดับ ${item.Quality_level}</span></div>
        <strong class="${item.trend.className}" title="${item.trend.label}">${item.trend.symbol}</strong>
      </div>
    `).join("");

    elements.execRankingPanel.innerHTML = `
      <div class="exec-rank-group"><h4>Top 5 ผลดีที่สุด</h4><div class="exec-rank-list">${renderRows(best)}</div></div>
      <div class="exec-rank-group"><h4>Bottom 5 ควรเร่งดำเนินการ</h4><div class="exec-rank-list">${renderRows(bottom, true)}</div></div>
    `;
  }

  function renderExecutiveCriticalIndicators(filteredRecords, frequencyMap) {
    const total = filteredRecords.length || 1;
    const topItems = indicatorReference
      .map((item) => ({ ...item, count: frequencyMap[item.feature_column] || 0 }))
      .filter((item) => item.count > 0)
      .sort((a, b) => b.count - a.count || a.item_code.localeCompare(b.item_code, "th"))
      .slice(0, 8);

    if (!topItems.length) {
      elements.execCriticalIndicators.innerHTML = emptyState("ยังไม่พบตัวชี้วัดที่เป็นปัญหาในตัวกรองนี้");
      return;
    }

    const maxCount = Math.max(...topItems.map((item) => item.count), 1);
    elements.execCriticalIndicators.innerHTML = topItems.map((item, index) => `
      <div class="exec-bar-row">
        <div class="exec-bar-label" title="${item.indicator_text}">${index + 1}. ${item.item_code}</div>
        <div class="exec-bar-track">
          <div class="exec-bar-fill" style="width:${(item.count / maxCount) * 100}%; background:linear-gradient(90deg,#d94e4e,#f2a93b)"></div>
        </div>
        <div class="exec-bar-value">${formatNumber(item.count)} (${formatPercent((item.count / total) * 100)})</div>
      </div>
    `).join("");
  }

  function renderExecutiveComposition(filteredRecords) {
    const total = filteredRecords.length;
    if (!total) {
      elements.execCompositionChart.innerHTML = emptyState(NO_DATA_MESSAGE);
      return;
    }
    const colors = { A: "#2fa66a", B: "#4371ad", C: "#f2a93b", D: "#d94e4e" };
    const counts = qualityLevels.map((level) => ({
      level,
      count: filteredRecords.filter((record) => record.Quality_level === level).length,
      color: colors[level],
    }));
    let start = 0;
    const segments = counts.map((item) => {
      const width = (item.count / total) * 100;
      const left = start;
      start += width;
      return `<span style="left:${left}%;width:${width}%;background:${item.color}"></span>`;
    }).join("");
    const legend = counts.map((item) => `
      <div class="exec-legend-row">
        <span class="exec-dot" style="background:${item.color}"></span>
        <div><b>ระดับ ${item.level}</b><span>${formatPercent((item.count / total) * 100)}</span></div>
        <strong>${formatNumber(item.count)}</strong>
      </div>
    `).join("");
    elements.execCompositionChart.innerHTML = `
      <div class="stacked-bar">${segments}</div>
      <div class="exec-legend">${legend}</div>
    `;
  }

  function renderExecutiveGeo(filteredRecords) {
    if (!filteredRecords.length) {
      elements.execGeoAnalysis.innerHTML = emptyState(NO_DATA_MESSAGE);
      return;
    }
    const geoField = ["province", "จังหวัด", "region", "ภูมิภาค"].find((key) =>
      filteredRecords.some((record) => record[key] && String(record[key]).trim() !== "")
    );
    if (geoField) {
      elements.execGeoMode.textContent = "Map-ready";
    } else {
      elements.execGeoMode.textContent = "Priority Group Ranked Bar";
    }
    const groupInfo = geoField ? { key: geoField, label: geoField } : getPolicyGroupField(filteredRecords);
    const grouped = Object.values(filteredRecords.reduce((acc, record) => {
      const key = String(record[groupInfo.key] || record.center_code_anon || "ไม่ระบุ").trim() || "ไม่ระบุ";
      if (!acc[key]) acc[key] = { key, total: 0, d: 0, improveSum: 0 };
      acc[key].total += 1;
      acc[key].improveSum += record.ImproveScore;
      if (record.Quality_level === "D") acc[key].d += 1;
      return acc;
    }, {}))
      .map((item) => ({ ...item, avgScore: item.improveSum / item.total, dRate: (item.d / item.total) * 100 }))
      .sort((a, b) => b.avgScore - a.avgScore || b.dRate - a.dRate)
      .slice(0, 8);
    const maxScore = Math.max(...grouped.map((item) => item.avgScore), 1);
    elements.execGeoAnalysis.innerHTML = grouped.map((item) => `
      <div class="exec-bar-row">
        <div class="exec-bar-label" title="${item.key}">${item.key}</div>
        <div class="exec-bar-track">
          <div class="exec-bar-fill" style="width:${(item.avgScore / maxScore) * 100}%; background:linear-gradient(90deg,#f2a93b,#d94e4e)"></div>
        </div>
        <div class="exec-bar-value">เฉลี่ย ${item.avgScore.toFixed(1)}</div>
      </div>
    `).join("");
  }

  function renderExecutiveValidation(filteredRecords) {
    const keys = filteredRecords.map((record) => record.record_key).filter(Boolean);
    const duplicateCount = keys.length - new Set(keys).size;
    const invalidScore = filteredRecords.filter((record) => !Number.isInteger(record.ImproveScore) || record.ImproveScore < 0 || record.ImproveScore > indicatorColumns.length).length;
    const invalidQuality = filteredRecords.filter((record) => !qualityLevels.includes(record.Quality_level)).length;
    const missingKey = filteredRecords.filter((record) => !record.record_key || !record.center_code_anon).length;
    const yearList = [...new Set(filteredRecords.map((record) => record.assessment_year_be))].sort((a, b) => a - b);
    const checks = [
      { label: "จำนวน Record", detail: `${formatNumber(filteredRecords.length)} ระเบียนในตัวกรอง`, status: "ok" },
      { label: "Duplicate Records", detail: `${formatNumber(duplicateCount)} ระเบียนซ้ำจาก record_key`, status: duplicateCount ? "warn" : "ok" },
      { label: "Data Type / Outliers", detail: `${formatNumber(invalidScore)} ค่า ImproveScore ผิดช่วงหรือไม่เป็นจำนวนเต็ม`, status: invalidScore ? "issue" : "ok" },
      { label: "Quality Level", detail: `${formatNumber(invalidQuality)} ระเบียนนอกกลุ่ม A/B/C/D`, status: invalidQuality ? "issue" : "ok" },
      { label: "Missing Key", detail: `${formatNumber(missingKey)} ระเบียนไม่มี record_key หรือ center_code_anon`, status: missingKey ? "warn" : "ok" },
      { label: "ปีของข้อมูล", detail: yearList.length ? `${yearList[0]}-${yearList[yearList.length - 1]}` : "ไม่พบปีในตัวกรอง", status: yearList.length ? "ok" : "warn" },
    ];
    elements.execValidationList.innerHTML = checks.map((check) => `
      <div class="validation-item">
        <span class="validation-badge ${check.status}">${check.status === "ok" ? "OK" : check.status === "warn" ? "CHECK" : "ISSUE"}</span>
        <div><b>${check.label}</b><span>${check.detail}</span></div>
        <span></span>
      </div>
    `).join("");
  }

  function renderExecutiveAlerts(filteredRecords, frequencyMap) {
    const total = filteredRecords.length || 1;
    const yearRows = getYearRows(filteredRecords);
    const current = yearRows[yearRows.length - 1];
    const previous = yearRows[yearRows.length - 2];
    const dCount = filteredRecords.filter((record) => record.Quality_level === "D").length;
    const cCount = filteredRecords.filter((record) => record.Quality_level === "C").length;
    const topIssue = indicatorReference
      .map((item) => ({ ...item, count: frequencyMap[item.feature_column] || 0 }))
      .sort((a, b) => b.count - a.count || a.item_code.localeCompare(b.item_code, "th"))[0];
    const alerts = [];
    if (dCount > 0) {
      alerts.push({ level: "critical", label: "Critical", text: `พบระดับ D จำนวน ${formatNumber(dCount)} ระเบียน (${formatPercent((dCount / total) * 100)}) ควรจัดลำดับช่วยเหลือก่อน` });
    }
    if (cCount > 0) {
      alerts.push({ level: "warning", label: "Warning", text: `มีกลุ่มเฝ้าระวังระดับ C จำนวน ${formatNumber(cCount)} ระเบียน (${formatPercent((cCount / total) * 100)})` });
    }
    if (topIssue && topIssue.count > 0) {
      alerts.push({ level: "warning", label: "Warning", text: `ตัวชี้วัด ${topIssue.item_code} พบมากที่สุด ${formatNumber(topIssue.count)} ครั้ง (${formatPercent((topIssue.count / total) * 100)})` });
      const actionTheme = getActionTheme(topIssue);
      alerts.push({ level: "warning", label: "Action", text: `ข้อเสนอเชิงนโยบาย: ${actionTheme.title} โดยเน้น ${actionTheme.detail}` });
    }
    if (current && previous) {
      const diff = current.passRate - previous.passRate;
      alerts.push({
        level: diff >= 0 ? "positive" : "critical",
        label: diff >= 0 ? "Positive" : "Critical",
        text: `อัตราผ่านปี ${current.year} ${diff >= 0 ? "เพิ่มขึ้น" : "ลดลง"} ${Math.abs(diff).toFixed(1)} จุดจากปี ${previous.year}`,
      });
    }
    const passCount = filteredRecords.filter((record) => record.target_label === PASS_LABEL).length;
    const passRate = (passCount / total) * 100;
    if (passRate >= PASS_TARGET_RATE) {
      alerts.push({ level: "positive", label: "Positive", text: `อัตราผ่าน ${formatPercent(passRate)} สูงกว่า benchmark ${PASS_TARGET_RATE}% ภายใต้ตัวกรองปัจจุบัน` });
    } else {
      alerts.push({ level: "warning", label: "Warning", text: `อัตราผ่าน ${formatPercent(passRate)} ต่ำกว่า benchmark ${PASS_TARGET_RATE}% อยู่ ${(PASS_TARGET_RATE - passRate).toFixed(1)} จุด` });
    }
    elements.execAlerts.innerHTML = alerts.slice(0, 6).map((alert) => `
      <div class="exec-alert-item ${alert.level}">
        <span class="exec-alert-badge ${alert.level}">${alert.label}</span>
        <span>${alert.text}</span>
        <span></span>
      </div>
    `).join("");
  }

  function renderExecutiveDashboard(filteredRecords, frequencyMap) {
    if (!elements.execStatusDonut) return;
    renderExecutiveStatus(filteredRecords);
    renderExecutiveTrend(filteredRecords);
    renderExecutiveCategory(filteredRecords);
    renderExecutiveRankings(filteredRecords);
    renderExecutiveCriticalIndicators(filteredRecords, frequencyMap);
    renderExecutiveComposition(filteredRecords);
    renderExecutiveGeo(filteredRecords);
    renderExecutiveAlerts(filteredRecords, frequencyMap);
    renderExecutiveValidation(filteredRecords);
  }

  function renderRankingChart(filteredRecords) {
    if (filteredRecords.length === 0) {
      elements.rankingChart.innerHTML = emptyState(NO_DATA_MESSAGE);
      return;
    }

    const grouped = Object.values(
      filteredRecords.reduce((acc, record) => {
        if (!acc[record.center_code_anon]) {
          acc[record.center_code_anon] = { center_code_anon: record.center_code_anon, ImproveScore: 0 };
        }
        acc[record.center_code_anon].ImproveScore += record.ImproveScore;
        return acc;
      }, {})
    )
      .sort((a, b) => b.ImproveScore - a.ImproveScore || a.center_code_anon.localeCompare(b.center_code_anon))
      .slice(0, 10);

    const maxValue = Math.max(...grouped.map((item) => item.ImproveScore), 1);
    const width = 860;
    const leftPad = 156;
    const rightPad = 76;
    const barHeight = 20;
    const gap = 18;
    const chartWidth = width - leftPad - rightPad;
    const height = grouped.length * (barHeight + gap) + 34;
    const gridColor = getCssVar("--line", "#d7e9ee");
    const teal = getCssVar("--teal", "#20c4c8");
    const tealDeep = getCssVar("--teal-deep", "#0e9097");

    const gridLines = Array.from({ length: 4 }, (_, index) => {
      const value = ((index + 1) / 4) * maxValue;
      const x = leftPad + (value / maxValue) * chartWidth;
      return `
        <line x1="${x}" y1="6" x2="${x}" y2="${height - 18}" stroke="${gridColor}" stroke-dasharray="5 6"></line>
        <text x="${x}" y="${height - 2}" text-anchor="middle" font-size="11" fill="#6f8aa0">${Math.round(value)}</text>
      `;
    }).join("");

    const bars = grouped
      .map((item, index) => {
        const y = index * (barHeight + gap) + 10;
        const barWidth = (item.ImproveScore / maxValue) * chartWidth;
        return `
          <text x="0" y="${y + 15}" font-size="13" fill="#5b778b" font-weight="700">${item.center_code_anon}</text>
          <rect x="${leftPad}" y="${y}" width="${chartWidth}" height="${barHeight}" rx="10" fill="#eef9f8"></rect>
          <rect x="${leftPad}" y="${y}" width="${barWidth}" height="${barHeight}" rx="10" fill="url(#rankBarGradient)"></rect>
          <circle cx="${leftPad + barWidth}" cy="${y + barHeight / 2}" r="4.5" fill="${tealDeep}"></circle>
          <text x="${leftPad + barWidth + 12}" y="${y + 15}" font-size="13" fill="#253346" font-weight="800">${formatNumber(item.ImproveScore)}</text>
        `;
      })
      .join("");

    elements.rankingChart.innerHTML = `
      <svg class="chart-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="อันดับสถานพัฒนา">
        <defs>
          <linearGradient id="rankBarGradient" x1="0%" x2="100%" y1="0%" y2="0%">
            <stop offset="0%" stop-color="${teal}"></stop>
            <stop offset="100%" stop-color="${tealDeep}"></stop>
          </linearGradient>
        </defs>
        ${gridLines}
        ${bars}
      </svg>
    `;
  }

  function renderYearlyChart(filteredRecords) {
    const yearRows = years
      .filter((year) => state.selectedYears.has(year))
      .map((year) => {
        const items = filteredRecords.filter((record) => record.assessment_year_be === year);
        const total = items.length;
        const passCount = items.filter((record) => record.target_label === PASS_LABEL).length;
        const dCount = items.filter((record) => record.Quality_level === "D").length;
        return {
          year,
          passRate: total === 0 ? 0 : (passCount / total) * 100,
          dRate: total === 0 ? 0 : (dCount / total) * 100,
        };
      });

    if (yearRows.length === 0) {
      elements.yearlyChart.innerHTML = emptyState(NO_YEAR_DATA_MESSAGE);
      return;
    }

    const width = 760;
    const height = 360;
    const leftPad = 48;
    const topPad = 24;
    const bottomPad = 42;
    const chartHeight = height - topPad - bottomPad;
    const plotWidth = width - leftPad - 24;
    const stepX = yearRows.length === 1 ? 0 : plotWidth / (yearRows.length - 1);
    const teal = getCssVar("--teal", "#20c4c8");
    const tealDeep = getCssVar("--teal-deep", "#0e9097");
    const danger = getCssVar("--level-d", "#ef5a5a");
    const dangerDeep = getCssVar("--level-d-deep", "#d63b4a");

    const axisLabels = [0, 25, 50, 75, 100]
      .map((value) => {
        const y = topPad + chartHeight - (value / 100) * chartHeight;
        return `
          <line x1="${leftPad}" x2="${width - 12}" y1="${y}" y2="${y}" stroke="#d7e9ee" stroke-dasharray="4 5"></line>
          <text x="${leftPad - 10}" y="${y + 4}" text-anchor="end" font-size="12" fill="#6f8aa0">${value}%</text>
        `;
      })
      .join("");

    const passPoints = yearRows.map((row, index) => ({
      x: leftPad + index * stepX,
      y: topPad + chartHeight - (row.passRate / 100) * chartHeight,
      year: row.year,
      value: row.passRate,
    }));
    const failPoints = yearRows.map((row, index) => ({
      x: leftPad + index * stepX,
      y: topPad + chartHeight - (row.dRate / 100) * chartHeight,
      year: row.year,
      value: row.dRate,
    }));

    const passPath = passPoints.map((point, index) => `${index === 0 ? "M" : "L"} ${point.x} ${point.y}`).join(" ");
    const failPath = failPoints.map((point, index) => `${index === 0 ? "M" : "L"} ${point.x} ${point.y}`).join(" ");
    const passArea = `${passPath} L ${passPoints[passPoints.length - 1].x} ${topPad + chartHeight} L ${passPoints[0].x} ${topPad + chartHeight} Z`;
    const failArea = `${failPath} L ${failPoints[failPoints.length - 1].x} ${topPad + chartHeight} L ${failPoints[0].x} ${topPad + chartHeight} Z`;

    const pointsMarkup = yearRows
      .map((row, index) => `
        <circle cx="${passPoints[index].x}" cy="${passPoints[index].y}" r="5.5" fill="${tealDeep}" stroke="#ffffff" stroke-width="3"></circle>
        <circle cx="${failPoints[index].x}" cy="${failPoints[index].y}" r="5.5" fill="${dangerDeep}" stroke="#ffffff" stroke-width="3"></circle>
        <text x="${passPoints[index].x}" y="${passPoints[index].y - 12}" text-anchor="middle" font-size="11" fill="${tealDeep}" font-weight="700">${row.passRate.toFixed(1)}</text>
        <text x="${failPoints[index].x}" y="${failPoints[index].y - 12}" text-anchor="middle" font-size="11" fill="${dangerDeep}" font-weight="700">${row.dRate.toFixed(1)}</text>
        <text x="${passPoints[index].x}" y="${height - 12}" text-anchor="middle" font-size="12" fill="#5b778b" font-weight="700">${row.year}</text>
      `)
      .join("");

    elements.yearlyChart.innerHTML = `
      <svg class="chart-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="แนวโน้มรายปี">
        <defs>
          <linearGradient id="passAreaGradient" x1="0%" x2="0%" y1="0%" y2="100%">
            <stop offset="0%" stop-color="${teal}" stop-opacity="0.28"></stop>
            <stop offset="100%" stop-color="${teal}" stop-opacity="0.03"></stop>
          </linearGradient>
          <linearGradient id="failAreaGradient" x1="0%" x2="0%" y1="0%" y2="100%">
            <stop offset="0%" stop-color="${danger}" stop-opacity="0.18"></stop>
            <stop offset="100%" stop-color="${danger}" stop-opacity="0.03"></stop>
          </linearGradient>
        </defs>
        ${axisLabels}
        <path d="${failArea}" fill="url(#failAreaGradient)"></path>
        <path d="${passArea}" fill="url(#passAreaGradient)"></path>
        <path d="${passPath}" fill="none" stroke="${tealDeep}" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"></path>
        <path d="${failPath}" fill="none" stroke="${dangerDeep}" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"></path>
        ${pointsMarkup}
        <g transform="translate(${width - 190}, 10)">
          <rect x="0" y="0" width="14" height="14" rx="4" fill="${tealDeep}"></rect>
          <text x="22" y="12" font-size="12" fill="#5b778b">อัตราผ่าน</text>
          <rect x="100" y="0" width="14" height="14" rx="4" fill="${dangerDeep}"></rect>
          <text x="122" y="12" font-size="12" fill="#5b778b">สัดส่วนระดับ D</text>
        </g>
      </svg>
    `;
  }

  function renderSchoolTimeline(filteredRecords) {
    const availableSchools = schoolOptions.filter((option) =>
      filteredRecords.some((record) => record.center_code_anon === option.center_code_anon)
    );

    if (!availableSchools.length) {
      elements.selectedSchoolName.textContent = "-";
      elements.selectedSchoolCode.textContent = "-";
      elements.schoolTimelineChart.innerHTML = emptyState(NO_DATA_MESSAGE);
      elements.schoolTimelineSummary.innerHTML = "";
      return;
    }

    if (!availableSchools.some((option) => option.center_code_anon === state.selectedCenterCode)) {
      state.selectedCenterCode = availableSchools[0].center_code_anon;
    }

    const selectedSchool = availableSchools.find((option) => option.center_code_anon === state.selectedCenterCode);
    const schoolRecords = filteredRecords
      .filter((record) => record.center_code_anon === state.selectedCenterCode)
      .sort((a, b) => a.assessment_year_be - b.assessment_year_be);

    elements.selectedSchoolName.textContent = selectedSchool.center_alias;
    elements.selectedSchoolCode.textContent = "ชื่อจริงถูกซ่อน";
    elements.schoolSearch.value = getSchoolDisplay(selectedSchool);

    if (!schoolRecords.length) {
      elements.schoolTimelineChart.innerHTML = emptyState(NO_YEAR_DATA_MESSAGE);
      elements.schoolTimelineSummary.innerHTML = "";
      return;
    }

    const width = 920;
    const height = 240;
    const leftPad = 80;
    const rightPad = 40;
    const topPad = 36;
    const bottomPad = 70;
    const lineY = 120;
    const plotWidth = width - leftPad - rightPad;
    const stepX = years.length === 1 ? 0 : plotWidth / (years.length - 1);
    const yearMap = Object.fromEntries(schoolRecords.map((record) => [record.assessment_year_be, record]));
    const gradeFill = {
      A: [getCssVar("--level-a", "#20c4c8"), getCssVar("--level-a-deep", "#0e9097")],
      B: [getCssVar("--level-b", "#67b7ff"), getCssVar("--level-b-deep", "#4a8fed")],
      C: [getCssVar("--level-c", "#f7bf56"), getCssVar("--level-c-deep", "#ef9a3c")],
      D: [getCssVar("--level-d", "#ef5a5a"), getCssVar("--level-d-deep", "#d63b4a")],
    };

    const connector = years
      .map((year, index) => {
        if (index === years.length - 1) return "";
        const x1 = leftPad + index * stepX;
        const x2 = leftPad + (index + 1) * stepX;
        return `<line x1="${x1}" y1="${lineY}" x2="${x2}" y2="${lineY}" stroke="#d7e9ee" stroke-width="6" stroke-linecap="round"></line>`;
      })
      .join("");

    const points = years
      .map((year, index) => {
        const x = leftPad + index * stepX;
        const record = yearMap[year];
        if (!record) {
          return `
            <circle cx="${x}" cy="${lineY}" r="12" fill="#eef9f8" stroke="#d7e9ee" stroke-width="3"></circle>
            <text x="${x}" y="${lineY + 38}" text-anchor="middle" font-size="12" fill="#6f8aa0" font-weight="700">${year}</text>
            <text x="${x}" y="${lineY + 58}" text-anchor="middle" font-size="11" fill="#9ab4c3">-</text>
          `;
        }

        const [startColor, endColor] = gradeFill[record.Quality_level] || gradeFill.A;
        return `
          <defs>
            <linearGradient id="grade-${year}" x1="0%" x2="100%" y1="0%" y2="100%">
              <stop offset="0%" stop-color="${startColor}"></stop>
              <stop offset="100%" stop-color="${endColor}"></stop>
            </linearGradient>
          </defs>
          <circle cx="${x}" cy="${lineY}" r="24" fill="url(#grade-${year})" stroke="#ffffff" stroke-width="5"></circle>
          <text x="${x}" y="${lineY + 6}" text-anchor="middle" font-size="20" fill="#ffffff" font-weight="800">${record.Quality_level}</text>
          <text x="${x}" y="${lineY - 42}" text-anchor="middle" font-size="11" fill="#6f8aa0">ผลการประเมิน</text>
          <text x="${x}" y="${lineY - 24}" text-anchor="middle" font-size="14" fill="#253346" font-weight="800">${record.Quality_level}</text>
          <text x="${x}" y="${lineY + 42}" text-anchor="middle" font-size="12" fill="#5b778b" font-weight="700">${year}</text>
          <text x="${x}" y="${lineY + 60}" text-anchor="middle" font-size="11" fill="#6f8aa0">${record.target_label}</text>
        `;
      })
      .join("");

    elements.schoolTimelineChart.innerHTML = `
      <svg class="chart-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="ผลการประเมินรายโรงเรียน">
        <rect x="${leftPad}" y="${lineY - 3}" width="${plotWidth}" height="6" rx="3" fill="#eef9f8"></rect>
        ${connector}
        ${points}
      </svg>
    `;

    elements.schoolTimelineSummary.innerHTML = schoolRecords
      .map(
        (record) => `
          <div class="timeline-pill">
            <span>${record.assessment_year_be}</span>
            <span class="timeline-grade grade-${record.Quality_level.toLowerCase()}">${record.Quality_level}</span>
            <strong>${record.target_label}</strong>
          </div>
        `
      )
      .join("");
  }

  function renderTopIndicators(filteredRecords, frequencyMap) {
    const topItems = indicatorReference
      .map((item) => ({ ...item, count: frequencyMap[item.feature_column] || 0 }))
      .sort((a, b) => b.count - a.count || a.item_code.localeCompare(b.item_code, "th"))
      .slice(0, 6);

    elements.topIndicators.innerHTML = "";
    const maxCount = Math.max(...topItems.map((item) => item.count), 1);
    const total = filteredRecords.length || 1;
    topItems.forEach((item, index) => {
      const tile = document.createElement("div");
      const barWidth = (item.count / maxCount) * 100;
      const share = (item.count / total) * 100;
      tile.className = `indicator-tile urgency-${index + 1}`;
      tile.innerHTML = `
        <div class="indicator-rank">Top ${index + 1}</div>
        <div class="indicator-bar-content">
          <div class="indicator-title-row">
            <strong class="indicator-title">${item.item_code}</strong>
            <span class="indicator-meta">พบ ${formatNumber(item.count)} ระเบียน (${formatPercent(share)})</span>
          </div>
          <div class="indicator-meta">${truncate(item.indicator_text, 150)}</div>
          <div class="indicator-bar-track">
            <span class="indicator-bar-fill" style="width:${barWidth}%"></span>
          </div>
        </div>
        <div class="indicator-count">${formatNumber(item.count)}</div>
      `;
      elements.topIndicators.appendChild(tile);
    });
  }

  function renderTable(filteredRecords, frequencyMap) {
    const rows = filteredRecords
      .slice()
      .sort((a, b) => b.ImproveScore - a.ImproveScore || a.center_code_anon.localeCompare(b.center_code_anon))
      .slice(0, 40);

    elements.trackingTableBody.innerHTML = "";
    if (!rows.length) {
      elements.trackingTableBody.innerHTML = `
        <tr><td colspan="5" style="text-align:center;color:#6f8aa0;">${NO_DATA_MESSAGE}</td></tr>
      `;
      return;
    }

    rows.forEach((row) => {
      const topIndicator = getTopIndicatorForRecord(row, frequencyMap);
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td><strong>${row.center_code_anon}</strong><br><span style="color:#6f8aa0;">ซ่อนชื่อจริงเพื่อคุ้มครองข้อมูลส่วนบุคคล</span><br><span style="color:#6f8aa0;">ปี ${row.assessment_year_be}</span></td>
        <td><span class="pill ${row.target_label === PASS_LABEL ? "pass" : "fail"}">${row.target_label}</span></td>
        <td>${formatNumber(row.ImproveScore)}</td>
        <td><span class="pill level-${row.Quality_level.toLowerCase()}">${row.Quality_level}</span></td>
        <td><strong>${topIndicator.label}</strong><br><span style="color:#6f8aa0;">${truncate(topIndicator.detail, 96)}</span></td>
      `;
      elements.trackingTableBody.appendChild(tr);
    });
  }

  function render() {
    renderFilters();
    const filteredRecords = getFilteredRecords();
    const frequencyMap = computeIndicatorFrequency(filteredRecords);
    renderExecutiveDashboard(filteredRecords, frequencyMap);
    renderKpis(filteredRecords);
    renderPolicyAction(filteredRecords, frequencyMap);
    renderRankingChart(filteredRecords);
    renderYearlyChart(filteredRecords);
    renderSchoolTimeline(filteredRecords);
    renderTopIndicators(filteredRecords, frequencyMap);
    renderTable(filteredRecords, frequencyMap);
  }

  function wireEvents() {
    elements.resetYears.addEventListener("click", () => {
      state.selectedYears = new Set(years);
      render();
    });

    elements.resetQuality.addEventListener("click", () => {
      state.selectedQuality = new Set(qualityLevels);
      render();
    });

    const applySchoolSearch = () => {
      const selected = findSchoolOptionByInput(elements.schoolSearch.value);
      if (selected) {
        state.selectedCenterCode = selected.center_code_anon;
        render();
      }
    };

    elements.schoolSearch.addEventListener("change", applySchoolSearch);
    elements.schoolSearch.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        applySchoolSearch();
      }
    });
  }

  function renderHeaderMeta() {
    const totalCenters = new Set(records.map((record) => record.center_code_anon).filter(Boolean)).size;
    const assessedRecords = records.filter((record) => record.target_label).length;
    const passCount = records.filter((record) => record.target_label === PASS_LABEL).length;
    const failCount = assessedRecords - passCount;
    const passRate = assessedRecords === 0 ? 0 : (passCount / assessedRecords) * 100;
    const failRate = assessedRecords === 0 ? 0 : (failCount / assessedRecords) * 100;
    const qualityD = records.filter((record) => record.Quality_level === "D").length;

    elements.sourceSheet.textContent = dashboardData.source.sheet;
    elements.totalRecords.textContent = `${formatNumber(totalCenters)} โรงเรียน`;
    elements.generatedAt.textContent = dashboardData.source.generated_at;
    elements.headerPassRate.textContent = `ผ่าน ${formatPercent(passRate)}`;
    elements.headerFailRate.textContent = `ไม่ผ่าน ${formatPercent(failRate)} | ${formatNumber(failCount)} ระเบียน`;
    elements.headerQualityD.textContent = `${formatNumber(qualityD)} ระเบียน`;
    elements.headerAssessedRecords.textContent = `ประเมินทั้งหมด ${formatNumber(assessedRecords)} ระเบียน`;
  }

  renderSchoolSearchOptions();
  renderHeaderMeta();
  wireEvents();
  render();
})();
