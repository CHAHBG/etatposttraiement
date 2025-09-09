/* dashboard_charts.js */
(function () {
  const palette = {
    primary: '#1e40af',
    blue: '#3b82f6',
    lightBlue: '#93c5fd',
    success: '#059669',
    green: '#10b981',
    warning: '#f59e0b',
    danger: '#ef4444',
    purple: '#8b5cf6',
    gray: '#6b7280'
  };

  const charts = {};
  let localCommunes = window.communesData || [];
  let collectionData = window.collectionData || [];
  let projectionData = window.projectionData || [];
  let yieldsData = window.yieldsData || [];
  let teamData = window.teamData || [];
  let nicadProjectionData = window.nicadProjectionData || [];
  let publicDisplayData = window.publicDisplayData || [];
  let ctasfProjectionData = window.ctasfProjectionData || [];

  // Defensive plugin registration: some Chart.js plugins auto-register, but
  // depending on load order and remote builds the plugin global may exist but
  // not be registered with Chart. Try common global names and register them
  // if present to avoid runtime mismatches (safety: failures are non-fatal).
  (function tryRegisterChartPlugins() {
    const candidates = [
      'ChartDataLabels', // common for chartjs-plugin-datalabels
      'chartjsPluginDatalabels',
      'ChartSankey', // possible exports for sankey plugin
      'chartjsChartSankey',
      'ChartSankeyPlugin'
    ];
    candidates.forEach(name => {
      try {
        const plugin = typeof window !== 'undefined' ? window[name] : undefined;
        if (plugin) {
          try { Chart.register(plugin); console.info('Registered Chart plugin:', name); } catch (e) { console.warn('Chart.register failed for', name, e); }
        }
      } catch (e) {}
    });
  })();

  function toNum(v) {
    if (v === null || v === undefined) return 0;
    if (typeof v === 'number') return v;
    const n = Number(String(v).replace(/[^0-9.-]+/g, ''));
    return Number.isFinite(n) ? n : 0;
  }

  function safeDestroyCanvas(target) {
    try {
      const canvas = target instanceof HTMLCanvasElement ? target : target.canvas || null;
      if (canvas) {
        const existing = Chart.getChart(canvas);
        if (existing) existing.destroy();
      }
    } catch (e) {}
  }

  function renderStatusChart() {
    const canvas = document.getElementById('statusChart');
    if (!canvas) return;
    safeDestroyCanvas(canvas);
    const data = {
      labels: ['Collected', 'Surveyed', 'Validated', 'Rejected'],
      datasets: [{
        label: 'Parcel Status',
        data: [
          localCommunes.reduce((s, c) => s + toNum(c.collected), 0),
          localCommunes.reduce((s, c) => s + toNum(c.surveyed), 0),
          localCommunes.reduce((s, c) => s + toNum(c.validated), 0),
          localCommunes.reduce((s, c) => s + toNum(c.rejected), 0)
        ],
        backgroundColor: [palette.blue, palette.green, palette.success, palette.danger],
        borderColor: ['#fff'],
        borderWidth: 1
      }]
    };

    if (charts.statusChart) charts.statusChart.destroy();
    charts.statusChart = new Chart(canvas, {
      type: 'bar',
      data: data,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          datalabels: {
            color: '#fff',
            font: { weight: 'bold', size: 12 },
            formatter: value => value.toLocaleString(),
            anchor: 'end',
            align: 'top'
          },
          tooltip: {
            backgroundColor: palette.primary,
            titleFont: { size: 14 },
            bodyFont: { size: 12 },
            callbacks: {
              label: ctx => `${ctx.label}: ${ctx.raw.toLocaleString()} parcels`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: true,
            title: { display: true, text: 'Parcels', color: palette.gray, font: { size: 12 } },
            grid: { color: palette.gray + '33' }
          },
          x: {
            title: { display: true, text: 'Status', color: palette.gray, font: { size: 12 } }
          }
        }
      }
    });
  }

  function renderCommunesChart() {
    const canvas = document.getElementById('communesChart');
    if (!canvas) return;
    safeDestroyCanvas(canvas);
    const topCommunes = localCommunes.sort((a, b) => b.collected - a.collected).slice(0, 5);
    const data = {
      labels: topCommunes.map(c => c.commune),
      datasets: [{
        label: 'Parcels Collected',
        data: topCommunes.map(c => c.collected),
        backgroundColor: palette.primary,
        borderRadius: 6,
        borderWidth: 1,
        borderColor: '#fff'
      }]
    };

    if (charts.communesChart) charts.communesChart.destroy();
    charts.communesChart = new Chart(canvas, {
      type: 'bar',
      data: data,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          datalabels: {
            color: '#fff',
            font: { weight: 'bold', size: 12 },
            anchor: 'end',
            align: 'top',
            formatter: value => value.toLocaleString()
          },
          tooltip: {
            backgroundColor: palette.primary,
            titleFont: { size: 14 },
            bodyFont: { size: 12 },
            callbacks: {
              label: ctx => `${ctx.raw.toLocaleString()} parcels`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: true,
            title: { display: true, text: 'Parcels', color: palette.gray, font: { size: 12 } },
            grid: { color: palette.gray + '33' }
          },
          x: {
            title: { display: true, text: 'Communes', color: palette.gray, font: { size: 12 } }
          }
        }
      }
    });
  }

  function renderRegionChart() {
    const canvas = document.getElementById('regionChart');
    if (!canvas) return;
    safeDestroyCanvas(canvas);
    const regions = {};
    localCommunes.forEach(c => {
      if (c.region) {
        regions[c.region] = (regions[c.region] || 0) + c.collected;
      }
    });
    const regionData = Object.entries(regions).sort((a, b) => b[1] - a[1]).slice(0, 5);

    if (charts.regionChart) charts.regionChart.destroy();
    charts.regionChart = new Chart(canvas, {
      type: 'pie',
      data: {
        labels: regionData.map(([region]) => region),
        datasets: [{
          label: 'Parcels by Region',
          data: regionData.map(([, count]) => count),
          backgroundColor: [palette.blue, palette.green, palette.purple, palette.warning, palette.danger],
          borderColor: '#fff',
          borderWidth: 1
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'bottom',
            labels: { color: palette.gray, font: { size: 12 } }
          },
          datalabels: {
            color: '#fff',
            font: { weight: 'bold', size: 12 },
            formatter: value => value.toLocaleString()
          },
          tooltip: {
            backgroundColor: palette.primary,
            titleFont: { size: 14 },
            bodyFont: { size: 12 },
            callbacks: {
              label: ctx => `${ctx.label}: ${ctx.raw.toLocaleString()} parcels`
            }
          }
        }
      }
    });
  }

  function renderCollectionPhaseChart() {
    const canvas = document.getElementById('collectionPhaseChart');
    if (!canvas) return;
    safeDestroyCanvas(canvas);
    const phases = {};
    collectionData.forEach(c => {
      if (c.phase) {
        phases[c.phase] = (phases[c.phase] || 0) + c.collected;
      }
    });
    const phaseData = Object.entries(phases);

    if (charts.collectionPhaseChart) charts.collectionPhaseChart.destroy();
    charts.collectionPhaseChart = new Chart(canvas, {
      type: 'pie',
      data: {
        labels: phaseData.map(([phase]) => phase),
        datasets: [{
          label: 'Collection by Phase',
          data: phaseData.map(([, count]) => count),
          backgroundColor: [palette.blue, palette.green, palette.purple],
          borderColor: '#fff',
          borderWidth: 1
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'bottom',
            labels: { color: palette.gray, font: { size: 12 } }
          },
          datalabels: {
            color: '#fff',
            font: { weight: 'bold', size: 12 },
            formatter: value => value.toLocaleString()
          },
          tooltip: {
            backgroundColor: palette.primary,
            titleFont: { size: 14 },
            bodyFont: { size: 12 }
          }
        }
      }
    });
  }

  function renderToolUtilizationChart() {
    const canvas = document.getElementById('toolUtilizationChart');
    if (!canvas) return;
    safeDestroyCanvas(canvas);
    // Placeholder: Assuming tool data comes from Team Deployment sheet
    const tools = ['Tool A', 'Tool B', 'Tool C'];
    const toolUsage = tools.map((_, i) => teamData.reduce((s, t) => s + (i % 2 ? t.assignedParcels : t.assignedParcels / 2), 0));
    const data = {
      labels: tools,
      datasets: [{
        label: 'Tool Utilization',
        data: toolUsage.length ? toolUsage : [60, 25, 15], // Fallback
        backgroundColor: [palette.blue, palette.green, palette.purple],
        borderColor: '#fff',
        borderWidth: 1
      }]
    };

    if (charts.toolUtilizationChart) charts.toolUtilizationChart.destroy();
    charts.toolUtilizationChart = new Chart(canvas, {
      type: 'doughnut',
      data: data,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'bottom',
            labels: { color: palette.gray, font: { size: 12 } }
          },
          datalabels: {
            color: '#fff',
            font: { weight: 'bold', size: 12 },
            formatter: value => `${Math.round(value / toolUsage.reduce((s, v) => s + v, 0) * 100)}%`
          },
          tooltip: {
            backgroundColor: palette.primary,
            titleFont: { size: 14 },
            bodyFont: { size: 12 },
            callbacks: {
              label: ctx => `${ctx.label}: ${Math.round(ctx.raw / toolUsage.reduce((s, v) => s + v, 0) * 100)}%`
            }
          }
        }
      }
    });
  }

  function renderDailyCollectionTrend() {
    const canvas = document.getElementById('dailyCollectionTrendChart');
    if (!canvas) return;
    safeDestroyCanvas(canvas);
    const dailyData = collectionData.reduce((acc, c) => {
      if (c.date) {
        acc[c.date] = (acc[c.date] || 0) + c.collected;
      }
      return acc;
    }, {});
    const labels = Object.keys(dailyData).slice(-7); // Last 7 days
    const data = {
      labels: labels,
      datasets: [{
        label: 'Daily Collections',
        data: labels.map(date => dailyData[date] || 0),
        borderColor: palette.primary,
        backgroundColor: palette.lightBlue + '66',
        fill: true,
        tension: 0.4,
        pointRadius: 4,
        pointHoverRadius: 6
      }]
    };

    if (charts.dailyCollectionTrendChart) charts.dailyCollectionTrendChart.destroy();
    charts.dailyCollectionTrendChart = new Chart(canvas, {
      type: 'line',
      data: data,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: palette.primary,
            titleFont: { size: 14 },
            bodyFont: { size: 12 },
            callbacks: {
              label: ctx => `${ctx.raw.toLocaleString()} parcels`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: true,
            title: { display: true, text: 'Parcels', color: palette.gray, font: { size: 12 } },
            grid: { color: palette.gray + '33' }
          },
          x: {
            title: { display: true, text: 'Date', color: palette.gray, font: { size: 12 } }
          }
        }
      }
    });
  }

  function renderMultiMonthProjectionChart() {
    const canvas = document.getElementById('multiMonthProjectionChart');
    if (!canvas) return;
    safeDestroyCanvas(canvas);
    const data = {
      labels: projectionData.map(p => p.month),
      datasets: [{
        label: 'Projected Parcels',
        data: projectionData.map(p => p.projectedParcels),
        backgroundColor: palette.blue,
        borderRadius: 6,
        borderWidth: 1,
        borderColor: '#fff'
      }]
    };

    if (charts.multiMonthProjectionChart) charts.multiMonthProjectionChart.destroy();
    charts.multiMonthProjectionChart = new Chart(canvas, {
      type: 'bar',
      data: data,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          datalabels: {
            color: '#fff',
            font: { weight: 'bold', size: 12 },
            anchor: 'end',
            align: 'top',
            formatter: value => value.toLocaleString()
          },
          tooltip: {
            backgroundColor: palette.primary,
            titleFont: { size: 14 },
            bodyFont: { size: 12 },
            callbacks: {
              label: ctx => `${ctx.raw.toLocaleString()} parcels`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: true,
            title: { display: true, text: 'Parcels', color: palette.gray, font: { size: 12 } },
            grid: { color: palette.gray + '33' }
          },
          x: {
            title: { display: true, text: 'Month', color: palette.gray, font: { size: 12 } }
          }
        }
      }
    });
  }

  function renderNicadAssignmentChart() {
    const canvas = document.getElementById('nicadAssignmentChart');
    if (!canvas) return;
    safeDestroyCanvas(canvas);
    const data = {
      labels: nicadProjectionData.map(p => p.week),
      datasets: [{
        label: 'NICAD Assignments',
        data: nicadProjectionData.map(p => p.assignments),
        borderColor: palette.green,
        backgroundColor: palette.green + '33',
        fill: true,
        tension: 0.4,
        pointRadius: 4,
        pointHoverRadius: 6
      }]
    };

    if (charts.nicadAssignmentChart) charts.nicadAssignmentChart.destroy();
    charts.nicadAssignmentChart = new Chart(canvas, {
      type: 'line',
      data: data,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: palette.primary,
            titleFont: { size: 14 },
            bodyFont: { size: 12 },
            callbacks: {
              label: ctx => `${ctx.raw.toLocaleString()} assignments`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: true,
            title: { display: true, text: 'Assignments', color: palette.gray, font: { size: 12 } },
            grid: { color: palette.gray + '33' }
          },
          x: {
            title: { display: true, text: 'Week', color: palette.gray, font: { size: 12 } }
          }
        }
      }
    });
  }

  function renderCTASFCompletionChart() {
    const canvas = document.getElementById('ctasfCompletionChart');
    if (!canvas) return;
    safeDestroyCanvas(canvas);
    const data = {
      labels: ctasfProjectionData.map(p => p.week),
      datasets: [{
        label: 'CTASF Completion',
        data: ctasfProjectionData.map(p => p.completion),
        backgroundColor: palette.purple,
        borderRadius: 6,
        borderWidth: 1,
        borderColor: '#fff'
      }]
    };

    if (charts.ctasfCompletionChart) charts.ctasfCompletionChart.destroy();
    charts.ctasfCompletionChart = new Chart(canvas, {
      type: 'bar',
      data: data,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          datalabels: {
            color: '#fff',
            font: { weight: 'bold', size: 12 },
            anchor: 'end',
            align: 'top',
            formatter: value => `${value.toFixed(1)}%`
          },
          tooltip: {
            backgroundColor: palette.primary,
            titleFont: { size: 14 },
            bodyFont: { size: 12 },
            callbacks: {
              label: ctx => `${ctx.raw.toFixed(1)}% completion`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: true,
            max: 100,
            title: { display: true, text: 'Completion (%)', color: palette.gray, font: { size: 12 } },
            grid: { color: palette.gray + '33' }
          },
          x: {
            title: { display: true, text: 'Week', color: palette.gray, font: { size: 12 } }
          }
        }
      }
    });
  }

  function renderPublicDisplayChart() {
    const canvas = document.getElementById('publicDisplayChart');
    if (!canvas) return;
    safeDestroyCanvas(canvas);
    const data = {
      labels: publicDisplayData.map(p => p.week),
      datasets: [{
        label: 'Public Display Progress',
        data: publicDisplayData.map(p => p.progress),
        backgroundColor: palette.success,
        borderRadius: 6,
        borderWidth: 1,
        borderColor: '#fff'
      }]
    };

    if (charts.publicDisplayChart) charts.publicDisplayChart.destroy();
    charts.publicDisplayChart = new Chart(canvas, {
      type: 'bar',
      data: data,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          datalabels: {
            color: '#fff',
            font: { weight: 'bold', size: 12 },
            anchor: 'end',
            align: 'top',
            formatter: value => `${value.toFixed(1)}%`
          },
          tooltip: {
            backgroundColor: palette.primary,
            titleFont: { size: 14 },
            bodyFont: { size: 12 },
            callbacks: {
              label: ctx => `${ctx.raw.toFixed(1)}% progress`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: true,
            max: 100,
            title: { display: true, text: 'Progress (%)', color: palette.gray, font: { size: 12 } },
            grid: { color: palette.gray + '33' }
          },
          x: {
            title: { display: true, text: 'Week', color: palette.gray, font: { size: 12 } }
          }
        }
      }
    });
  }

  function renderErrorRateChart() {
    const canvas = document.getElementById('errorRateChart');
    if (!canvas) return;
    safeDestroyCanvas(canvas);
    const errorRates = localCommunes.map(c => ({
      commune: c.commune,
      errorRate: c.rejected && c.collected ? (c.rejected / c.collected * 100).toFixed(1) : 0
    })).sort((a, b) => b.errorRate - a.errorRate).slice(0, 5);

    if (charts.errorRateChart) charts.errorRateChart.destroy();
    charts.errorRateChart = new Chart(canvas, {
      type: 'bar',
      data: {
        labels: errorRates.map(c => c.commune),
        datasets: [{
          label: 'Error Rate (%)',
          data: errorRates.map(c => c.errorRate),
          backgroundColor: palette.danger,
          borderRadius: 6,
          borderWidth: 1,
          borderColor: '#fff'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          datalabels: {
            color: '#fff',
            font: { weight: 'bold', size: 12 },
            anchor: 'end',
            align: 'top',
            formatter: value => `${value}%`
          },
          tooltip: {
            backgroundColor: palette.primary,
            titleFont: { size: 14 },
            bodyFont: { size: 12 },
            callbacks: {
              label: ctx => `${ctx.raw}% error rate`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: true,
            max: 100,
            title: { display: true, text: 'Error Rate (%)', color: palette.gray, font: { size: 12 } },
            grid: { color: palette.gray + '33' }
          },
          x: {
            title: { display: true, text: 'Commune', color: palette.gray, font: { size: 12 } }
          }
        }
      }
    });
  }

  function renderDuplicateRemovalChart() {
    const canvas = document.getElementById('duplicateRemovalChart');
    if (!canvas) return;
    safeDestroyCanvas(canvas);
    const duplicateRates = localCommunes.map(c => ({
      commune: c.commune,
      rate: c.duplicateRemovalRatePct || 0
    })).sort((a, b) => b.rate - a.rate).slice(0, 5);

    if (charts.duplicateRemovalChart) charts.duplicateRemovalChart.destroy();
    charts.duplicateRemovalChart = new Chart(canvas, {
      type: 'bar',
      data: {
        labels: duplicateRates.map(c => c.commune),
        datasets: [{
          label: 'Duplicate Removal Rate (%)',
          data: duplicateRates.map(c => c.rate),
          backgroundColor: palette.green,
          borderRadius: 6,
          borderWidth: 1,
          borderColor: '#fff'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          datalabels: {
            color: '#fff',
            font: { weight: 'bold', size: 12 },
            anchor: 'end',
            align: 'top',
            formatter: value => `${value}%`
          },
          tooltip: {
            backgroundColor: palette.primary,
            titleFont: { size: 14 },
            bodyFont: { size: 12 },
            callbacks: {
              label: ctx => `${ctx.raw}% duplicate removal`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: true,
            max: 100,
            title: { display: true, text: 'Removal Rate (%)', color: palette.gray, font: { size: 12 } },
            grid: { color: palette.gray + '33' }
          },
          x: {
            title: { display: true, text: 'Commune', color: palette.gray, font: { size: 12 } }
          }
        }
      }
    });
  }

  function renderUrmValidationChart() {
    const canvas = document.getElementById('urmValidationChart');
    if (!canvas) return;
    safeDestroyCanvas(canvas);
    const validationData = localCommunes.map(c => ({
      commune: c.commune,
      validated: c.validated || 0
    })).sort((a, b) => b.validated - a.validated).slice(0, 5);

    if (charts.urmValidationChart) charts.urmValidationChart.destroy();
    charts.urmValidationChart = new Chart(canvas, {
      type: 'bar',
      data: {
        labels: validationData.map(c => c.commune),
        datasets: [{
          label: 'URM Validated Parcels',
          data: validationData.map(c => c.validated),
          backgroundColor: palette.success,
          borderRadius: 6,
          borderWidth: 1,
          borderColor: '#fff'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          datalabels: {
            color: '#fff',
            font: { weight: 'bold', size: 12 },
            anchor: 'end',
            align: 'top',
            formatter: value => value.toLocaleString()
          },
          tooltip: {
            backgroundColor: palette.primary,
            titleFont: { size: 14 },
            bodyFont: { size: 12 },
            callbacks: {
              label: ctx => `${ctx.raw.toLocaleString()} parcels`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: true,
            title: { display: true, text: 'Validated Parcels', color: palette.gray, font: { size: 12 } },
            grid: { color: palette.gray + '33' }
          },
          x: {
            title: { display: true, text: 'Commune', color: palette.gray, font: { size: 12 } }
          }
        }
      }
    });
  }

  function renderQualityScoreChart() {
    const canvas = document.getElementById('qualityScoreChart');
    if (!canvas) return;
    safeDestroyCanvas(canvas);
    const qualityScores = localCommunes.map(c => ({
      commune: c.commune,
      score: c.validated && c.collected ? (c.validated / c.collected * 100).toFixed(1) : 0
    })).sort((a, b) => b.score - a.score).slice(0, 5);

    if (charts.qualityScoreChart) charts.qualityScoreChart.destroy();
    charts.qualityScoreChart = new Chart(canvas, {
      type: 'bar',
      data: {
        labels: qualityScores.map(c => c.commune),
        datasets: [{
          label: 'Quality Score (%)',
          data: qualityScores.map(c => c.score),
          backgroundColor: palette.purple,
          borderRadius: 6,
          borderWidth: 1,
          borderColor: '#fff'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          datalabels: {
            color: '#fff',
            font: { weight: 'bold', size: 12 },
            anchor: 'end',
            align: 'top',
            formatter: value => `${value}%`
          },
          tooltip: {
            backgroundColor: palette.primary,
            titleFont: { size: 14 },
            bodyFont: { size: 12 },
            callbacks: {
              label: ctx => `${ctx.raw}% quality score`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: true,
            max: 100,
            title: { display: true, text: 'Quality Score (%)', color: palette.gray, font: { size: 12 } },
            grid: { color: palette.gray + '33' }
          },
          x: {
            title: { display: true, text: 'Commune', color: palette.gray, font: { size: 12 } }
          }
        }
      }
    });
  }

  function renderProcessFlowChart() {
    const canvas = document.getElementById('processFlowChart');
    if (!canvas) return;
    safeDestroyCanvas(canvas);
    const flowData = [
      { from: 'Collected', to: 'Surveyed', value: localCommunes.reduce((s, c) => s + toNum(c.surveyed), 0) },
      { from: 'Surveyed', to: 'Validated', value: localCommunes.reduce((s, c) => s + toNum(c.validated), 0) },
      { from: 'Surveyed', to: 'Rejected', value: localCommunes.reduce((s, c) => s + toNum(c.rejected), 0) }
    ];

    if (charts.processFlowChart) charts.processFlowChart.destroy();
    try {
      charts.processFlowChart = new Chart(canvas, {
        type: 'sankey',
        data: {
          datasets: [{
            data: flowData,
            colorFrom: palette.blue,
            colorTo: palette.green,
            colorMode: 'gradient',
            labels: {
              'Collected': 'Parcels Collected',
              'Surveyed': 'Parcels Surveyed',
              'Validated': 'Parcels Validated',
              'Rejected': 'Parcels Rejected'
            }
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: {
              backgroundColor: palette.primary,
              titleFont: { size: 14 },
              bodyFont: { size: 12 },
              callbacks: {
                label: ctx => `${ctx.raw.value.toLocaleString()} parcels`
              }
            }
          }
        }
      });
    } catch (err) {
      console.error('Failed to render Sankey/process flow chart (plugin missing or incompatible):', err);
      // Fallback: render a simple bar chart with the same high-level totals
      try {
        const totals = {
          collected: localCommunes.reduce((s, c) => s + toNum(c.collected), 0),
          surveyed: localCommunes.reduce((s, c) => s + toNum(c.surveyed), 0),
          validated: localCommunes.reduce((s, c) => s + toNum(c.validated), 0),
          rejected: localCommunes.reduce((s, c) => s + toNum(c.rejected), 0)
        };
        const labels = ['Collected', 'Surveyed', 'Validated', 'Rejected'];
        const data = [totals.collected, totals.surveyed, totals.validated, totals.rejected];
        // Always ensure any existing Chart instance tied to this canvas is destroyed before creating a new one.
        try { safeDestroyCanvas(canvas); } catch (e) { /* ignore */ }
        if (charts.processFlowChart) { try { charts.processFlowChart.destroy(); } catch (e) {} }
        charts.processFlowChart = new Chart(canvas, {
          type: 'bar',
          data: { labels: labels, datasets: [{ label: 'Process Flow (fallback)', data: data, backgroundColor: [palette.blue, palette.lightBlue, palette.green, palette.danger] }] },
          options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true } } }
        });
      } catch (err2) {
        console.error('Fallback process flow chart also failed:', err2);
      }
    }
  }

  function renderGeomaticianPerformanceChart() {
    const canvas = document.getElementById('geomaticianPerformanceChart');
    if (!canvas) return;
    safeDestroyCanvas(canvas);
    const performance = {};
    teamData.forEach(t => {
      if (t.geomatician) {
        performance[t.geomatician] = (performance[t.geomatician] || 0) + t.assignedParcels;
      }
    });
    const performanceData = Object.entries(performance).sort((a, b) => b[1] - a[1]).slice(0, 5);

    if (charts.geomaticianPerformanceChart) charts.geomaticianPerformanceChart.destroy();
    charts.geomaticianPerformanceChart = new Chart(canvas, {
      type: 'bar',
      data: {
        labels: performanceData.map(([name]) => name),
        datasets: [{
          label: 'Assigned Parcels',
          data: performanceData.map(([, count]) => count),
          backgroundColor: palette.blue,
          borderRadius: 6,
          borderWidth: 1,
          borderColor: '#fff'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          datalabels: {
            color: '#fff',
            font: { weight: 'bold', size: 12 },
            anchor: 'end',
            align: 'top',
            formatter: value => value.toLocaleString()
          },
          tooltip: {
            backgroundColor: palette.primary,
            titleFont: { size: 14 },
            bodyFont: { size: 12 },
            callbacks: {
              label: ctx => `${ctx.raw.toLocaleString()} parcels`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: true,
            title: { display: true, text: 'Assigned Parcels', color: palette.gray, font: { size: 12 } },
            grid: { color: palette.gray + '33' }
          },
          x: {
            title: { display: true, text: 'Geomatician', color: palette.gray, font: { size: 12 } }
          }
        }
      }
    });
  }

  function renderTeamEfficiencyChart() {
    const canvas = document.getElementById('teamEfficiencyChart');
    if (!canvas) return;
    safeDestroyCanvas(canvas);
    const efficiency = {};
    localCommunes.forEach(c => {
      if (c.geomatician && c.validated > 0) {
        efficiency[c.geomatician] = (efficiency[c.geomatician] || 0) + c.validated;
      }
    });
    const efficiencyData = Object.entries(efficiency).sort((a, b) => b[1] - a[1]).slice(0, 5);

    if (charts.teamEfficiencyChart) charts.teamEfficiencyChart.destroy();
    charts.teamEfficiencyChart = new Chart(canvas, {
      type: 'bar',
      data: {
        labels: efficiencyData.map(([name]) => name),
        datasets: [{
          label: 'Validated Parcels',
          data: efficiencyData.map(([, count]) => count),
          backgroundColor: palette.purple,
          borderRadius: 6,
          borderWidth: 1,
          borderColor: '#fff'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          datalabels: {
            color: '#fff',
            font: { weight: 'bold', size: 12 },
            anchor: 'end',
            align: 'top',
            formatter: value => value.toLocaleString()
          },
          tooltip: {
            backgroundColor: palette.primary,
            titleFont: { size: 14 },
            bodyFont: { size: 12 },
            callbacks: {
              label: ctx => `${ctx.raw.toLocaleString()} parcels`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: true,
            title: { display: true, text: 'Validated Parcels', color: palette.gray, font: { size: 12 } },
            grid: { color: palette.gray + '33' }
          },
          x: {
            title: { display: true, text: 'Geomatician', color: palette.gray, font: { size: 12 } }
          }
        }
      }
    });
  }

  function renderBottleneckIndicators() {
    const container = document.getElementById('bottleneckIndicators');
    if (!container) return;
    container.innerHTML = '';
    const bottlenecks = [
      {
        name: 'Data Collection',
        status: localCommunes.some(c => c.collected === 0) ? 'red' : 'green',
        description: 'Checks if any commune has zero collected parcels'
      },
      {
        name: 'Validation',
        status: localCommunes.some(c => c.validated / c.collected < 0.5) ? 'yellow' : 'green',
        description: 'Validation rate below 50% in any commune'
      },
      {
        name: 'Rejection',
        status: localCommunes.some(c => c.rejected / c.collected > 0.2) ? 'red' : 'green',
        description: 'Rejection rate above 20% in any commune'
      },
      {
        name: 'NICAD Assignment',
        status: nicadProjectionData.some(p => p.assignments === 0) ? 'yellow' : 'green',
        description: 'Checks for zero NICAD assignments in any week'
      }
    ];

    bottlenecks.forEach(b => {
      const div = document.createElement('div');
      div.className = 'bottleneck-indicator';
      div.setAttribute('aria-label', `${b.name} status: ${b.status}, ${b.description}`);
      div.innerHTML = `
        <div class="light ${b.status}"></div>
        <span>${b.name}</span>
        <small>${b.description}</small>
      `;
      container.appendChild(div);
    });
  }

  function renderMethodologyChecklist() {
    const container = document.getElementById('methodologyChecklist');
    if (!container) return;
    container.innerHTML = '';
    const methodology = window._procasef_methodology_notes || [
      { section: 'Data Collection', completed: true },
      { section: 'Validation', completed: false },
      { section: 'Public Display', completed: false }
    ];

    methodology.forEach(m => {
      const div = document.createElement('div');
      div.className = `check-item ${m.completed ? 'completed' : ''}`;
      div.setAttribute('aria-label', `${m.section} status: ${m.completed ? 'Completed' : 'Incomplete'}`);
      div.innerHTML = `
        <i class="fas fa-${m.completed ? 'check-circle' : 'times-circle'}"></i>
        <span>${m.section}</span>
      `;
      container.appendChild(div);
    });
  }

  function refreshAll() {
    Object.values(charts).forEach(chart => chart?.destroy());
    renderStatusChart();
    renderCommunesChart();
    renderRegionChart();
    renderCollectionPhaseChart();
    renderToolUtilizationChart();
    renderDailyCollectionTrend();
    renderMultiMonthProjectionChart();
    renderNicadAssignmentChart();
    renderCTASFCompletionChart();
    renderPublicDisplayChart();
    renderErrorRateChart();
    renderDuplicateRemovalChart();
    renderUrmValidationChart();
    renderQualityScoreChart();
    renderProcessFlowChart();
    renderGeomaticianPerformanceChart();
    renderTeamEfficiencyChart();
    renderBottleneckIndicators();
    renderMethodologyChecklist();
  }

  window.addEventListener('procasef:data:loaded', (e) => {
    const sheets = e.detail && e.detail.data ? e.detail.data : {};
    if (Array.isArray(sheets['Commune Status'])) {
      localCommunes = sheets['Commune Status'].map(r => ({
        commune: r['Commune'] || '',
        region: r['Region'] || '',
        collected: toNum(r['Collected Parcels (No Duplicates)']),
        surveyed: toNum(r['Surveyed Parcels']),
        validated: toNum(r['URM Validated Parcels']),
        rejected: toNum(r['URM Rejected Parcels']),
        geomatician: r['Geomatician'] || '',
        duplicateRemovalRatePct: toNum(r['Duplicate Removal Rate (%)']),
        retained: toNum(r['Retained Parcels'])
      }));
    }
    if (Array.isArray(sheets['Data Collection'])) {
      collectionData = sheets['Data Collection'].map(r => ({
        phase: r['Phase'] || '',
        collected: toNum(r['Collected']),
        date: r['Date'] || ''
      }));
    }
    if (Array.isArray(sheets['Projection Collection'])) {
      projectionData = sheets['Projection Collection'].map(r => ({
        month: r['Month'] || '',
        projectedParcels: toNum(r['Projected Parcels'])
      }));
    }
    if (Array.isArray(sheets['Yields Projection'])) {
      yieldsData = sheets['Yields Projection'].map(r => ({
        commune: r['Commune'] || '',
        yieldEstimate: toNum(r['Yield Estimate'])
      }));
    }
    if (Array.isArray(sheets['Team Deployment'])) {
      teamData = sheets['Team Deployment'].map(r => ({
        geomatician: r['Geomatician'] || '',
        assignedParcels: toNum(r['Assigned Parcels']),
        startDate: r['Start Date'] || '',
        endDate: r['End Date'] || ''
      }));
    }
    if (Array.isArray(sheets['NICAD Projection'])) {
      nicadProjectionData = sheets['NICAD Projection'].map(r => ({
        week: r['Week'] || '',
        assignments: toNum(r['Assignments'])
      }));
    }
    if (Array.isArray(sheets['Public Display'])) {
      publicDisplayData = sheets['Public Display'].map(r => ({
        week: r['Week'] || '',
        progress: toNum(r['Progress %'])
      }));
    }
    if (Array.isArray(sheets['CTASF Projection'])) {
      ctasfProjectionData = sheets['CTASF Projection'].map(r => ({
        week: r['Week'] || '',
        completion: toNum(r['Completion %'])
      }));
    }
    setTimeout(refreshAll, 80);
  });

  window.addEventListener('procasef:communes:updated', (e) => {
    if (Array.isArray(e.detail.data)) {
      localCommunes = e.detail.data;
      setTimeout(refreshAll, 50);
    }
  });

  document.addEventListener('DOMContentLoaded', () => {
    setTimeout(refreshAll, 200);
    const observer = new IntersectionObserver((entries, observer) => {
      entries.forEach(entry => {
        if (entry.isIntersecting) {
          const canvas = entry.target.querySelector('canvas, [id="bottleneckIndicators"], [id="methodologyChecklist"]');
          if (canvas) {
            const id = canvas.id;
            if (id === 'statusChart') renderStatusChart();
            else if (id === 'communesChart') renderCommunesChart();
            else if (id === 'regionChart') renderRegionChart();
            else if (id === 'collectionPhaseChart') renderCollectionPhaseChart();
            else if (id === 'toolUtilizationChart') renderToolUtilizationChart();
            else if (id === 'dailyCollectionTrendChart') renderDailyCollectionTrend();
            else if (id === 'multiMonthProjectionChart') renderMultiMonthProjectionChart();
            else if (id === 'nicadAssignmentChart') renderNicadAssignmentChart();
            else if (id === 'ctasfCompletionChart') renderCTASFCompletionChart();
            else if (id === 'publicDisplayChart') renderPublicDisplayChart();
            else if (id === 'errorRateChart') renderErrorRateChart();
            else if (id === 'duplicateRemovalChart') renderDuplicateRemovalChart();
            else if (id === 'urmValidationChart') renderUrmValidationChart();
            else if (id === 'qualityScoreChart') renderQualityScoreChart();
            else if (id === 'processFlowChart') renderProcessFlowChart();
            else if (id === 'geomaticianPerformanceChart') renderGeomaticianPerformanceChart();
            else if (id === 'teamEfficiencyChart') renderTeamEfficiencyChart();
            else if (id === 'bottleneckIndicators') renderBottleneckIndicators();
            else if (id === 'methodologyChecklist') renderMethodologyChecklist();
            observer.unobserve(entry.target);
          }
        }
      });
    }, { threshold: 0.1 });
    document.querySelectorAll('.chart-container').forEach(container => observer.observe(container));
  });

  let resizeTimeout;
  window.addEventListener('resize', () => {
    clearTimeout(resizeTimeout);
    resizeTimeout = setTimeout(() => Object.values(charts).forEach(c => c?.resize()), 100);
  });
})();