(function () {
  const links = document.querySelectorAll('[data-view-link]');
  const views = document.querySelectorAll('.view[data-view]');
  const splashText = document.getElementById('transitionSplashText');

  const checkMissingOrBtn = document.getElementById('checkMissingOrBtn');
  const openGenerateFrecnoBtn = document.getElementById('openGenerateFrecnoBtn');
  const openFrecnoListSqlBtn = document.getElementById('openFrecnoListSqlBtn');
  const openGenerateAllScriptBtn = document.getElementById('openGenerateAllScriptBtn');
  const missingOrModal = document.getElementById('missingOrModal');
  const closeMissingOrModal = document.getElementById('closeMissingOrModal');
  const orFileInput = document.getElementById('orFileInput');
  const uploadOrFileBtn = document.getElementById('uploadOrFileBtn');
  const clearOrFileBtn = document.getElementById('clearOrFileBtn');
  const analyzeStatus = document.getElementById('analyzeStatus');

  const missingSummary = document.getElementById('missingSummary');
  const detectedCount = document.getElementById('detectedCount');
  const missingCount = document.getElementById('missingCount');
  const duplicateCount = document.getElementById('duplicateCount');
  const duplicateDetailMismatchCount = document.getElementById('duplicateDetailMismatchCount');
  const rollbackCount = document.getElementById('rollbackCount');
  const outOfOrderCount = document.getElementById('outOfOrderCount');
  const missingListOutput = document.getElementById('missingListOutput');
  const sqlActions = document.getElementById('sqlActions');
  const generateSelectSqlBtn = document.getElementById('generateSelectSqlBtn');
  const generateDeleteSqlBtn = document.getElementById('generateDeleteSqlBtn');
  const sqlResultOutput = document.getElementById('sqlResultOutput');
  const copyMissingSqlRow = document.getElementById('copyMissingSqlRow');
  const copyMissingSqlBtn = document.getElementById('copyMissingSqlBtn');
  const openAyalaCheckerBtn = document.getElementById('openAyalaCheckerBtn');
  const openXmlRenameBtn = document.getElementById('openXmlRenameBtn');
  const openReportDocxBtn = document.getElementById('openReportDocxBtn');
  const pfpPromptInput = document.getElementById('pfpPromptInput');
  const pfpApiKeyInput = document.getElementById('pfpApiKeyInput');
  const pfpFileInput = document.getElementById('pfpFileInput');
  const generatePfpStableBtn = document.getElementById('generatePfpStableBtn');
  const generatePfpBtn = document.getElementById('generatePfpBtn');
  const downloadPfpBtn = document.getElementById('downloadPfpBtn');
  const clearPfpBtn = document.getElementById('clearPfpBtn');
  const pfpStatus = document.getElementById('pfpStatus');
  const pfpCanvas = document.getElementById('pfpCanvas');
  const ayalaModal = document.getElementById('ayalaModal');
  const closeAyalaModal = document.getElementById('closeAyalaModal');
  const hourlyCsvInput = document.getElementById('hourlyCsvInput');
  const eodCsvInput = document.getElementById('eodCsvInput');
  const uploadAyalaBtn = document.getElementById('uploadAyalaBtn');
  const mergeAyalaBtn = document.getElementById('mergeAyalaBtn');
  const clearAyalaBtn = document.getElementById('clearAyalaBtn');
  const ayalaStatus = document.getElementById('ayalaStatus');
  const ayalaDiscrepancyActions = document.getElementById('ayalaDiscrepancyActions');
  const ayalaTxnDiscrepancyBtn = document.getElementById('ayalaTxnDiscrepancyBtn');
  const ayalaTotalDiscrepancyBtn = document.getElementById('ayalaTotalDiscrepancyBtn');
  const copyAyalaReportRow = document.getElementById('copyAyalaReportRow');
  const copyAyalaReportBtn = document.getElementById('copyAyalaReportBtn');
  const downloadAyalaCheckerBtn = document.getElementById('downloadAyalaCheckerBtn');
  const ayalaReportOutput = document.getElementById('ayalaReportOutput');
  const ayalaResultModal = document.getElementById('ayalaResultModal');
  const closeAyalaResultModal = document.getElementById('closeAyalaResultModal');
  const okAyalaResultBtn = document.getElementById('okAyalaResultBtn');
  const ayalaResultText = document.getElementById('ayalaResultText');
  const ayalaGuidelinesTabBtn = document.getElementById('ayalaGuidelinesTabBtn');
  const ayalaCheckerTabBtn = document.getElementById('ayalaCheckerTabBtn');
  const ayalaGuidelinesPanel = document.getElementById('ayalaGuidelinesPanel');
  const ayalaCheckerPanel = document.getElementById('ayalaCheckerPanel');
  const ayalaGuidelinesZipPath = 'files/Ayala guidelines.zip';
  const eodUploadCard = eodCsvInput ? eodCsvInput.closest('.ayala-upload-card') : null;

  const frecnoModal = document.getElementById('frecnoModal');
  const closeFrecnoModal = document.getElementById('closeFrecnoModal');
  const oldFrecnoInput = document.getElementById('oldFrecnoInput');
  const latestFrecnoInput = document.getElementById('latestFrecnoInput');
  const generateFrecnoBtn = document.getElementById('generateFrecnoBtn');
  const clearFrecnoBtn = document.getElementById('clearFrecnoBtn');
  const frecnoOutput = document.getElementById('frecnoOutput');
  const frecnoStatus = document.getElementById('frecnoStatus');
  const frecnoProgress = document.getElementById('frecnoProgress');
  const copyFrecnoCodeRow = document.getElementById('copyFrecnoCodeRow');
  const copyFrecnoCodeBtn = document.getElementById('copyFrecnoCodeBtn');
  const xmlRenameModal = document.getElementById('xmlRenameModal');
  const closeXmlRenameModal = document.getElementById('closeXmlRenameModal');
  const xmlRenameFileInput = document.getElementById('xmlRenameFileInput');
  const xmlTenantIdInput = document.getElementById('xmlTenantIdInput');
  const xmlKeyInput = document.getElementById('xmlKeyInput');
  const processXmlRenameBtn = document.getElementById('processXmlRenameBtn');
  const clearXmlRenameBtn = document.getElementById('clearXmlRenameBtn');
  const xmlRenameStatus = document.getElementById('xmlRenameStatus');
  const xmlRenameOutput = document.getElementById('xmlRenameOutput');
  const reportDocxModal = document.getElementById('reportDocxModal');
  const closeReportDocxModal = document.getElementById('closeReportDocxModal');
  const reportDocxTemplateInput = document.getElementById('reportDocxTemplateInput');
  const downloadReportTemplateBtn = document.getElementById('downloadReportTemplateBtn');
  const reportDocxStatus = document.getElementById('reportDocxStatus');
  const reportDocxOutput = document.getElementById('reportDocxOutput');
  const reportDateInput = document.getElementById('reportDateInput');
  const reportJofInput = document.getElementById('reportJofInput');
  const reportInvoiceInput = document.getElementById('reportInvoiceInput');
  const reportClientInput = document.getElementById('reportClientInput');
  const reportTsAssignedInput = document.getElementById('reportTsAssignedInput');
  const reportTitleInput = document.getElementById('reportTitleInput');
  const reportConcernInput = document.getElementById('reportConcernInput');
  const reportActivitiesEditor = document.getElementById('reportActivitiesEditor');
  const reportRootCauseInput = document.getElementById('reportRootCauseInput');
  const reportPreventiveInput = document.getElementById('reportPreventiveInput');
  const reportNextStepsInput = document.getElementById('reportNextStepsInput');
  const reportCurrentStatusInput = document.getElementById('reportCurrentStatusInput');
  const openActivitiesEditorBtn = document.getElementById('openActivitiesEditorBtn');
  const activitiesEditorState = document.getElementById('activitiesEditorState');
  const activitiesEditorDialog = document.getElementById('activitiesEditorDialog');
  const closeActivitiesEditorBtn = document.getElementById('closeActivitiesEditorBtn');
  const okActivitiesEditorBtn = document.getElementById('okActivitiesEditorBtn');
  const activitiesBulletBtn = document.getElementById('activitiesBulletBtn');
  const generateReportDocxBtn = document.getElementById('generateReportDocxBtn');
  const clearReportDocxBtn = document.getElementById('clearReportDocxBtn');
  const allScriptModal = document.getElementById('allScriptModal');
  const closeAllScriptModal = document.getElementById('closeAllScriptModal');
  const allScriptOutput = document.getElementById('allScriptOutput');
  const copyAllScriptBtn = document.getElementById('copyAllScriptBtn');
  const frecnoListSqlModal = document.getElementById('frecnoListSqlModal');
  const closeFrecnoListSqlModal = document.getElementById('closeFrecnoListSqlModal');
  const frecnoListSqlInput = document.getElementById('frecnoListSqlInput');
  const generateFrecnoListSelectBtn = document.getElementById('generateFrecnoListSelectBtn');
  const generateFrecnoListDeleteBtn = document.getElementById('generateFrecnoListDeleteBtn');
  const clearFrecnoListSqlBtn = document.getElementById('clearFrecnoListSqlBtn');
  const frecnoListSqlStatus = document.getElementById('frecnoListSqlStatus');
  const frecnoListSqlOutput = document.getElementById('frecnoListSqlOutput');
  const copyFrecnoListSqlRow = document.getElementById('copyFrecnoListSqlRow');
  const copyFrecnoListSqlBtn = document.getElementById('copyFrecnoListSqlBtn');

  let isTransitioning = false;
  let missingOrNumbers = [];
  let isAyalaDownloadInProgress = false;
  let ayalaMergedHourlyData = null;
  let ayalaMergedHourlySignature = '';
  let ayalaHourlyAnalyzed = false;
  let latestAyalaCheckerData = null;
  let latestAyalaTxnValidationReport = '';
  let latestAyalaTotalValidationReport = '';
  const REPORT_DOCX_TEMPLATE_PATH = 'files/report maker.docx';
  let latestPfpImage = null;
  let latestPfpBlobUrl = '';
  let latestPfpRemoteUrl = '';
  let latestPfpCanvasExportSafe = true;

  const ALL_SCRIPT_SECTIONS = [
    {
      title: 'CHECK missing sales',
      body: 'SELECT fzcounter, fsale_date, MIN(fdocument_no) AS fmin, MAX(fdocument_no) AS fmax, SUM(fgross) AS fgross, COUNT(fdocument_no) AS fdocument_no, MAX(fdocument_no) - MIN(fdocument_no) + 1 AS Expected_OR_count, count(*)-(max(fdocument_no)-min(fdocument_no)+1) as missing_OR FROM pos_sale WHERE fcompanyid="WPOS-22061582" and ftermid="0003" AND fsale_date>="20240701" AND fdocument_no <> 0 GROUP BY fsale_date, fzcounter',
    },
    {
      title: 'FULL MISSING SCRIPT',
      body: 'select fsale_date, fzcounter, min(fdocument_no) as min_fdoc, max(fdocument_no) as max_fdoc, count(*) as trxcnt, (max(fdocument_no) - min(fdocument_no) + 1) as expected, case when count(*) = (max(fdocument_no) - min(fdocument_no) + 1) then \'Y\' else \'N\' end as same, count(*) - (max(fdocument_no) - min(fdocument_no) + 1) as lacking, sum(fgross) as fgross from pos_sale where fcompanyid="STORE-17080382" AND ftermid="0001" and fsale_date >= \'20230901\' and fsale_date <= \'20241231\' and fdocument_no <> \'0\' group by fsale_date, fzcounter having (same = \'N\')',
    },
    {
      title: 'FOR EXCEL',
      body: '=IF(E8 - E7 > 1, IF(E8-E7 > 2, E7 + 1 & "-" & E8 - 1, E8 -1), "OK")',
    },
    {
      title: 'SELECT OR/DATE SCRIPT',
      body: 'select * from pos_sale where fcompanyid="AEON-17091282" and ftermid="0107" and fsale_date=\'20240316\' and fzcounter=\'2251\'\n\nselect * from pos_sale where fcompanyid="AEON-17091282" and ftermid="0107" and fsale_date=\'20240316\'and frecno=\'2251\'\n\nselect * from pos_sale where fcompanyid="AEON-17091282" and ftermid="0097" and frecno="58714"',
    },
    {
      title: 'TRANSFER TO ANOTHER TERMINAL',
      body: 'update pos_sale set ftermid=fterm_updated_by where fcompanyid="AEON-17091282" AND ftermid="0106" and frecno=\'101117\'',
    },
    {
      title: 'TRANFER TO ANOTHER DATE',
      body: 'update pos_sale set fsale_date="20220630" where fcompanyid="GT-20072286" AND ftermid="0013" and fsale_date="20230314" and fdocument_no="172369" and fzcounter="558"\n\nupdate pos_sale set fsale_date="20220630" where fcompanyid="GT-20072286" AND ftermid="0013" and fsale_date="20230314" and fdocument_no="172369" and frecno="558"',
    },
    {
      title: 'UPDATED CHANGE SET',
      body: 'update pos_sale set fzcounter=\'number na iset\' where fcompanyid=\'\' and ftermid=\'\' and fsale_date=\'\' and fzcounter=\'0\'',
    },
    {
      title: 'NO FCREATED',
      body: 'update pos_sale set fcreated_date=fposted_date WHERE fcompanyid= \'SRFDI-15100982\' AND ftermid = \'0029\' and fcreated_date=\'\'',
    },
    {
      title: 'SEQUENCE',
      body: 'update pos_sale set fdocument_no=fdocument_no-339 where fcompanyid="CBTL-14032883" AND ftermid="2025" and fsale_date="20231229" and fdocument_no between "440249" and "440545"\n\nupdate pos_sale set fdocument_no=fdocument_no+5 where fcompanyid="SRFDI-15100982" AND ftermid="0015" and fdocument_no between "807663" and "807768"\n\nupdate pos_sale set fdocument_no=fdocument_no+4 where fcompanyid="SRFDI-15100982" AND ftermid="0015" and fsale_date between "20231216" and "20240421" and fdocument_no!=\'0\'',
    },
    {
      title: 'Script for tomorrow',
      body: 'update pos_sale set fdocument_no=fdocument_no+339 where fcompanyid="CBTL-14032883" AND ftermid="2025" and fsale_date between "20231231" and "20240131" and fdocument_no!="0"',
    },
    {
      title: 'DB BROWSER DELETE AND REMAIN SELECTED ONLY',
      body: 'DELETE from pos_sale where frecno not in (\'\', \'\', \'\',);\n\nDELETE from pos_sale_payment where frecno not in (\'\', \'\', \'\',);\n\nDELETE from pos_sale_product where frecno not in (\'\', \'\', \'\',);\n\nDELETE from pos_reading;\n\nDELETE from pos_reading_summary',
    },
    {
      title: 'FRECNO UPDATE',
      body: 'UPDATE pos_sale_payment SET frecno = CASE frecno WHEN \'81342\' THEN \'90001\' WHEN \'81343\' THEN \'90002\' ELSE frecno END;',
    },
    {
      title: 'DB BROWSER UPDATE FTRANSMIT',
      body: 'update pos_sale set ftransmit="0000";\n\nupdate pos_sale_payment set ftransmit="0000";\n\nupdate pos_sale_product set ftransmit="0000";',
    },
    {
      title: 'FOR EXCEPTION',
      body: 'URL/appserv/app/batch/sys_get_exception.php',
    },
    {
      title: 'FOR BYPASS OF EXCEPTION ID',
      body: 'URL/appserv/app/batch/sys_bypass_exception.php',
    },
    {
      title: 'REREADING',
      body: 'http://app.alliancewebpos.net/appserv/app/batch/fix/recreate_reading.php?fcompanyid=WPOS-22022482&ftermid=0004&fsale_date=20230201&fzcounter=315&fend_date=20230216&fcreate_flag=1',
    },
    {
      title: 'HIDE',
      body: 'update pos_reading set fcompanyid=\'xTITAYS-14032882\' WHERE fcompanyid="TITAYS-14032882" and ftermid="0004" and fsale_date=\'20240812\'',
    },
    {
      title: 'MERGING',
      body: 'update pos_sale set fzcounter="2288" where fcompanyid="AFG-15042282" and ftermid="0011" and fsale_date="20250521" and fzcounter="0"',
    },
    {
      title: 'Checked the Duplicate fdocument_no',
      body: 'SELECT fdocument_no, COUNT(*) AS duplicate_count, MIN(fsale_date) AS first_sale_date, MAX(fsale_date) AS last_sale_date FROM pos_sale WHERE fcompanyid = \'AEON-17091282\' AND ftermid = \'0172\' AND fsale_date >= \'20251101\' AND fdocument_no <> 0 GROUP BY fdocument_no HAVING COUNT(*) > 1 ORDER BY fdocument_no;',
    },
  ];

  function escapeHtml(value) {
    return String(value || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function getPfpContext() {
    if (!pfpCanvas) return null;
    const ctx = pfpCanvas.getContext('2d');
    return ctx || null;
  }

  function drawPfpPlaceholder() {
    const ctx = getPfpContext();
    if (!ctx || !pfpCanvas) return;
    ctx.clearRect(0, 0, pfpCanvas.width, pfpCanvas.height);
    ctx.fillStyle = '#0b1325';
    ctx.fillRect(0, 0, pfpCanvas.width, pfpCanvas.height);
    ctx.strokeStyle = 'rgba(154, 191, 255, 0.45)';
    ctx.lineWidth = 10;
    ctx.strokeRect(5, 5, pfpCanvas.width - 10, pfpCanvas.height - 10);
    ctx.beginPath();
    ctx.strokeStyle = 'rgba(255, 225, 96, 0.65)';
    ctx.lineWidth = 8;
    ctx.arc(pfpCanvas.width / 2, pfpCanvas.height / 2, pfpCanvas.width * 0.42, 0, Math.PI * 2);
    ctx.stroke();
    ctx.fillStyle = '#9fd0ff';
    ctx.font = 'bold 52px Arial';
    ctx.textAlign = 'center';
    ctx.fillText('PFP', pfpCanvas.width / 2, pfpCanvas.height / 2 + 18);
  }

  function drawImageCoverToCanvas(img) {
    const ctx = getPfpContext();
    if (!ctx || !pfpCanvas || !img) return;
    const cw = pfpCanvas.width;
    const ch = pfpCanvas.height;
    const iw = img.naturalWidth || img.width;
    const ih = img.naturalHeight || img.height;
    if (!iw || !ih) return;
    const scale = Math.max(cw / iw, ch / ih);
    const sw = cw / scale;
    const sh = ch / scale;
    const sx = (iw - sw) / 2;
    const sy = (ih - sh) / 2;

    ctx.clearRect(0, 0, cw, ch);
    ctx.fillStyle = '#0b1325';
    ctx.fillRect(0, 0, cw, ch);
    ctx.drawImage(img, sx, sy, sw, sh, 0, 0, cw, ch);
  }

  function readImageFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const img = new Image();
        img.onload = () => resolve(img);
        img.onerror = () => reject(new Error('Invalid image file.'));
        img.src = String(reader.result || '');
      };
      reader.onerror = () => reject(new Error('Failed to read image.'));
      reader.readAsDataURL(file);
    });
  }

  function loadImageFromBlob(blob) {
    return new Promise((resolve, reject) => {
      const objectUrl = URL.createObjectURL(blob);
      const img = new Image();
      img.onload = () => {
        if (latestPfpBlobUrl) {
          URL.revokeObjectURL(latestPfpBlobUrl);
        }
        latestPfpBlobUrl = objectUrl;
        resolve(img);
      };
      img.onerror = () => {
        URL.revokeObjectURL(objectUrl);
        reject(new Error('Failed to load generated image.'));
      };
      img.src = objectUrl;
    });
  }

  function loadLocalLogoImage() {
    return new Promise((resolve, reject) => {
      const logo = new Image();
      logo.onload = () => resolve(logo);
      logo.onerror = () => reject(new Error('Logo not found.'));
      logo.src = 'css/Alliancelogodesign.jpeg';
    });
  }

  async function overlayAllianceLogoOnCanvas() {
    const ctx = getPfpContext();
    if (!ctx || !pfpCanvas) return;
    try {
      const logo = await loadLocalLogoImage();
      const cw = pfpCanvas.width;
      const ch = pfpCanvas.height;
      const ratio = (logo.naturalWidth || logo.width || 1) / (logo.naturalHeight || logo.height || 1);
      const logoWidth = cw * 0.14;
      const logoHeight = logoWidth / ratio;
      const x = (cw - logoWidth) / 2;
      const y = ch * 0.74;
      ctx.save();
      ctx.globalAlpha = 0.95;
      ctx.drawImage(logo, x, y, logoWidth, logoHeight);
      ctx.restore();
    } catch (error) {
      // Keep generation running even if logo is missing.
    }
  }

  async function generateImageFromPrompt(promptText) {
    const enforcedPrompt = `${String(promptText || '').trim()}, portrait profile picture, centered face, wearing plain black t-shirt, black shirt, clean background`;
    const encoded = encodeURIComponent(enforcedPrompt);
    const endpoints = [
      `https://image.pollinations.ai/prompt/${encoded}?width=1080&height=1080&nologo=true&model=flux&seed=${Date.now()}`,
      `https://image.pollinations.ai/prompt/${encoded}?width=1080&height=1080&nologo=true&seed=${Date.now()}`,
    ];
    // Final page fallback (manual save) if API-like endpoints are down.
    latestPfpRemoteUrl = `https://pollinations.ai/p/${encoded}`;

    let lastErr = null;
    for (let i = 0; i < endpoints.length; i += 1) {
      const url = endpoints[i];
      latestPfpRemoteUrl = url;
      try {
        const response = await fetch(url, { cache: 'no-store' });
        if (!response.ok) {
          throw new Error(`Generator request failed: HTTP ${response.status}`);
        }
        const blob = await response.blob();
        const img = await loadImageFromBlob(blob);
        return { img, corsSafe: true };
      } catch (error) {
        lastErr = error;
        // Fallback: try direct image loading per endpoint.
        try {
          const img = await new Promise((resolve, reject) => {
            const remoteImg = new Image();
            remoteImg.crossOrigin = 'anonymous';
            remoteImg.onload = () => resolve(remoteImg);
            remoteImg.onerror = () => reject(new Error('Online image load failed.'));
            remoteImg.src = url;
          });
          return { img, corsSafe: false };
        } catch (error2) {
          lastErr = error2;
        }
      }
    }
    throw lastErr || new Error('All prompt endpoints failed.');
  }

  async function generateImageFromOpenAI(apiKey, promptText) {
    const enforcedPrompt = `${String(promptText || '').trim()}, portrait profile picture, centered face, wearing plain black t-shirt, black shirt, clean background, realistic face, suitable for instagram profile`;
    const response = await fetch('https://api.openai.com/v1/images/generations', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model: 'gpt-image-1',
        prompt: enforcedPrompt,
        size: '1024x1024',
      }),
    });
    if (!response.ok) {
      const errText = await response.text().catch(() => '');
      throw new Error(`OpenAI image request failed: HTTP ${response.status} ${errText}`);
    }
    const data = await response.json();
    const first = data && Array.isArray(data.data) ? data.data[0] : null;
    if (!first) {
      throw new Error('OpenAI image response is empty.');
    }
    if (first.b64_json) {
      const dataUrl = `data:image/png;base64,${first.b64_json}`;
      const img = await new Promise((resolve, reject) => {
        const i = new Image();
        i.onload = () => resolve(i);
        i.onerror = () => reject(new Error('Failed to load OpenAI base64 image.'));
        i.src = dataUrl;
      });
      return { img, corsSafe: true };
    }
    if (first.url) {
      latestPfpRemoteUrl = first.url;
      const img = await new Promise((resolve, reject) => {
        const i = new Image();
        i.crossOrigin = 'anonymous';
        i.onload = () => resolve(i);
        i.onerror = () => reject(new Error('Failed to load OpenAI URL image.'));
        i.src = first.url;
      });
      return { img, corsSafe: false };
    }
    throw new Error('OpenAI image format not supported.');
  }

  function testCanvasExportSafety() {
    const ctx = getPfpContext();
    if (!ctx || !pfpCanvas) return false;
    try {
      ctx.getImageData(0, 0, 1, 1);
      return true;
    } catch (error) {
      return false;
    }
  }

  function downloadPfpCanvasPng() {
    if (!pfpCanvas) return;
    const url = pfpCanvas.toDataURL('image/png');
    const link = document.createElement('a');
    link.href = url;
    link.download = 'instagram_profile_1080.png';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  function openPfpRemoteUrlInNewTab() {
    if (!latestPfpRemoteUrl) return false;
    const link = document.createElement('a');
    link.href = latestPfpRemoteUrl;
    link.target = '_blank';
    link.rel = 'noopener';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    return true;
  }

  function buildAllScriptPlainText() {
    return ALL_SCRIPT_SECTIONS
      .map((section) => `${section.title}:\n${section.body}`)
      .join('\n\n');
  }

  function renderAllScriptOutput() {
    if (!allScriptOutput) {
      return;
    }
    const html = ALL_SCRIPT_SECTIONS
      .map((section) => (
        `<section class="all-script-section">`
        + `<p class="all-script-title"><strong>${escapeHtml(section.title)}</strong></p>`
        + `<pre class="all-script-code">${escapeHtml(section.body)}</pre>`
        + `</section>`
      ))
      .join('');
    allScriptOutput.innerHTML = html;
  }

  drawPfpPlaceholder();

  function parseAyalaRawPairsFromText(text) {
    return String(text || '')
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean)
      .map((line) => {
        const parts = line.split(',').map((part) => String(part || '').trim());
        if (parts.length < 2) {
          return null;
        }
        const key = parts[0];
        const values = parts.slice(1);
        if (!key) {
          return null;
        }
        return { key, values };
      })
      .filter(Boolean);
  }

  function buildAyalaRawGridFromPairs(pairs) {
    const rows = Array.isArray(pairs) ? pairs : [];
    if (rows.length === 0) {
      return { isMulti: false, terminals: [], rows: [] };
    }

    const maxColumns = rows.reduce((max, row) => Math.max(max, Array.isArray(row.values) ? row.values.length : 0), 0);
    if (maxColumns <= 0) {
      return { isMulti: false, terminals: [], rows: [] };
    }

    const terRow = rows.find((row) => String(row.key || '').trim().toUpperCase() === 'TER_NO');
    const terminals = (terRow && Array.isArray(terRow.values) ? terRow.values : [])
      .map((value) => String(value || '').trim())
      .filter(Boolean);
    const terminalLabels = maxColumns > 1
      ? Array.from({ length: maxColumns }, (_, idx) => (terminals[idx] && terminals[idx].trim() ? terminals[idx].trim() : `TER${idx + 1}`))
      : [];

    const summaryTotalFields = new Set(['GROSS_SLS', 'VAT_AMNT', 'VATABLE_SLS', 'REFUND_AMT']);
    const gridRows = rows.map((row) => {
      const values = Array.from({ length: maxColumns }, (_, idx) => String((row.values || [])[idx] ?? '').trim());
      let total = '';
      if (maxColumns > 1 && summaryTotalFields.has(String(row.key || '').trim().toUpperCase())) {
        total = formatAyalaAmount(values.reduce((sum, value) => sum + parseNumericAyala(value), 0));
      }
      return {
        field: String(row.key || '').trim(),
        values,
        total,
      };
    });

    return {
      isMulti: maxColumns > 1,
      terminals: terminalLabels,
      rows: gridRows,
    };
  }

  function buildAyalaEodPairsByTerminal(parsedData) {
    const records = parsedData && parsedData.records ? parsedData.records : [];
    if (!records || records.length === 0) {
      return { isMulti: false, terminals: [], rows: [] };
    }

    const terminalSet = [];
    const seenTerminals = new Set();
    records.forEach((record) => {
      const ter = String(record.TER_NO || '').trim() || '-';
      if (!seenTerminals.has(ter)) {
        seenTerminals.add(ter);
        terminalSet.push(ter);
      }
    });

    const headerOrder = (parsedData && parsedData.headers && parsedData.headers.length > 0)
      ? parsedData.headers
      : Object.keys(records[0] || {});
    const uniqueHeaders = [];
    const seenHeaders = new Set();
    headerOrder.forEach((header) => {
      const key = String(header || '').trim();
      if (key && !seenHeaders.has(key)) {
        seenHeaders.add(key);
        uniqueHeaders.push(key);
      }
    });

    const summaryTotalFields = new Set(['GROSS_SLS', 'VAT_AMNT', 'VATABLE_SLS', 'REFUND_AMT']);
    const rows = uniqueHeaders.map((field) => {
      const perTerminalValues = terminalSet.map((ter) => {
        const terRecords = records.filter((item) => String(item.TER_NO || '').trim() === ter);
        const values = terRecords
          .map((record) => String(record && Object.prototype.hasOwnProperty.call(record, field) ? record[field] : '').trim())
          .filter((value) => value !== '');
        if (values.length === 0) {
          return '';
        }
        const numericPattern = /^-?\d+(?:\.\d+)?$/;
        const normalized = values.map((value) => value.replace(/,/g, ''));
        const isNumeric = normalized.every((value) => numericPattern.test(value));
        if (isNumeric) {
          const sum = normalized.reduce((acc, value) => acc + (Number.parseFloat(value) || 0), 0);
          const hasDecimal = normalized.some((value) => value.includes('.'));
          return hasDecimal ? formatAyalaAmount(sum) : String(Math.round(sum));
        }
        return Array.from(new Set(values)).join(' | ');
      });

      let totalValue = '';
      if (terminalSet.length > 1 && summaryTotalFields.has(field)) {
        totalValue = formatAyalaAmount(
          perTerminalValues.reduce((acc, value) => acc + parseNumericAyala(value), 0)
        );
      }

      return {
        field,
        values: perTerminalValues,
        total: totalValue,
      };
    });

    return {
      isMulti: terminalSet.length > 1,
      terminals: terminalSet,
      rows,
    };
  }

  async function readAyalaRawPairs(file, parsedData) {
    if (!file) {
      return { isMulti: false, terminals: [], rows: [] };
    }
    const lowerName = String(file.name || '').toLowerCase();
    if (lowerName.endsWith('.csv') || lowerName.endsWith('.txt')) {
      const text = await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(typeof reader.result === 'string' ? reader.result : '');
        reader.onerror = () => reject(new Error('Failed to read EOD text file.'));
        reader.readAsText(file);
      });
      const pairs = parseAyalaRawPairsFromText(text);
      if (pairs.length > 0) {
        const rawGrid = buildAyalaRawGridFromPairs(pairs);
        if (rawGrid && rawGrid.rows && rawGrid.rows.length > 0) {
          return rawGrid;
        }
        const grouped = buildAyalaEodPairsByTerminal(parsedData);
        if (grouped && grouped.rows && grouped.rows.length > 0) {
          return grouped;
        }
        return {
          isMulti: false,
          terminals: [],
          rows: pairs.map((pair) => ({
            field: String(pair.key ?? ''),
            values: [String((pair.values || [])[0] ?? '')],
            total: '',
          })),
        };
      }
    }

    return buildAyalaEodPairsByTerminal(parsedData);
  }

  function getSplashLabel(viewName) {
    const labels = {
      home: 'Home',
      'missing-sales': 'Missing Sales',
      'csv-tools': 'Sales Validator',
      'xml-tools': 'XML Tools',
      'report-maker': 'Report Maker',
    };

    return labels[viewName] || 'Home';
  }

  function setActiveView(viewName, updateHash) {
    views.forEach((view) => {
      view.classList.toggle('active', view.dataset.view === viewName);
    });

    links.forEach((link) => {
      link.classList.toggle('active', link.classList.contains('nav-btn') && link.dataset.viewLink === viewName);
    });

    if (updateHash) {
      window.location.hash = viewName;
    }
  }

  function switchViewWithTransition(viewName) {
    const exists = Array.from(views).some((view) => view.dataset.view === viewName);
    if (!exists || isTransitioning) {
      return;
    }

    const activeView = document.querySelector('.view.active')?.dataset.view;
    if (activeView === viewName) {
      return;
    }

    isTransitioning = true;

    if (splashText) {
      splashText.textContent = getSplashLabel(viewName);
    }

    document.body.classList.add('is-switching');

    window.setTimeout(() => {
      setActiveView(viewName, true);
      document.body.classList.remove('is-switching');
      isTransitioning = false;
    }, 2000);
  }

  function chunkArray(items, size) {
    const chunks = [];
    for (let i = 0; i < items.length; i += size) {
      chunks.push(items.slice(i, i + size));
    }
    return chunks;
  }

  function formatSqlNumberList(numbers) {
    const quoted = numbers.map((value) => `"${value}"`);
    return chunkArray(quoted, 10).map((row) => row.join(',')).join('\n');
  }

  function isGrossHeaderCell(value) {
    const text = String(value || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
    if (!text) return false;
    if (text === 'gross') return true;
    if (text === 'grosssls') return true;
    if (text === 'grosssales') return true;
    if (text === 'totalgross') return true;
    if (text === 'grossamount') return true;
    return false;
  }

  function isCustomerHeaderCell(value) {
    const text = String(value || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
    if (!text) return false;
    if (text === 'customer') return true;
    if (text === 'customername') return true;
    if (text === 'clientname') return true;
    if (text === 'name') return true;
    return false;
  }

  function isDateHeaderCell(value) {
    const text = String(value || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
    if (!text) return false;
    if (text === 'date') return true;
    if (text === 'trndate') return true;
    if (text === 'saledate') return true;
    if (text === 'transactiondate') return true;
    return false;
  }

  function normalizeDuplicateCustomer(value) {
    return String(value || '')
      .toLowerCase()
      .replace(/\s+/g, ' ')
      .trim();
  }

  function normalizeDuplicateDate(value) {
    const raw = String(value || '').trim();
    if (!raw) return '';
    const ymd = raw.match(/^(\d{4})[-/](\d{2})[-/](\d{2})$/);
    if (ymd) return `${ymd[1]}-${ymd[2]}-${ymd[3]}`;
    const mdy = raw.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
    if (mdy) {
      return `${mdy[3]}-${mdy[1].padStart(2, '0')}-${mdy[2].padStart(2, '0')}`;
    }
    return raw.toLowerCase();
  }

  function isSameGrossAmount(a, b) {
    return Math.abs(roundAyalaAmount(parseNumericAyala(a)) - roundAyalaAmount(parseNumericAyala(b))) <= 0.009;
  }

  function analyzeSequence(orderedEntries) {
    const missing = [];
    const issues = [];
    const duplicates = [];
    const rollbacks = [];
    const duplicateDetailMismatches = [];
    const seenAtRow = new Map();
    const duplicateKeys = new Set();
    const entries = (orderedEntries || []).map((item, index) => {
      if (typeof item === 'number') {
        return {
          receipt: item,
          row: index + 1,
          grossRaw: '',
          gross: 0,
          customerRaw: '',
          customer: '',
          dateRaw: '',
          date: '',
        };
      }
      const receipt = Number.parseInt(String(item.receipt ?? '').trim(), 10);
      return {
        receipt,
        row: Number.isFinite(item.row) ? item.row : (index + 1),
        grossRaw: String(item.grossRaw ?? ''),
        gross: parseNumericAyala(item.grossRaw ?? item.gross ?? 0),
        customerRaw: String(item.customerRaw ?? ''),
        customer: normalizeDuplicateCustomer(item.customerRaw ?? item.customer ?? ''),
        dateRaw: String(item.dateRaw ?? ''),
        date: normalizeDuplicateDate(item.dateRaw ?? item.date ?? ''),
      };
    }).filter((item) => Number.isFinite(item.receipt));

    for (let i = 0; i < entries.length; i += 1) {
      const current = entries[i];
      const value = current.receipt;
      if (!seenAtRow.has(value)) {
        seenAtRow.set(value, current);
        continue;
      }

      const first = seenAtRow.get(value);
      const duplicateKey = `${value}|${first.row}|${current.row}`;
      if (!duplicateKeys.has(duplicateKey)) {
        duplicateKeys.add(duplicateKey);
        const hasGrossA = String(first.grossRaw || '').trim() !== '';
        const hasGrossB = String(current.grossRaw || '').trim() !== '';
        const grossSame = (() => {
          if (!hasGrossA && !hasGrossB) return true;
          if (hasGrossA !== hasGrossB) return false;
          return isSameGrossAmount(first.gross, current.gross);
        })();
        const grossMismatch = !grossSame;

        const hasCustomerA = String(first.customerRaw || '').trim() !== '';
        const hasCustomerB = String(current.customerRaw || '').trim() !== '';
        const customerSame = (() => {
          if (!hasCustomerA && !hasCustomerB) return true;
          if (hasCustomerA !== hasCustomerB) return false;
          return first.customer === current.customer;
        })();
        const customerMismatch = !customerSame;

        const hasDateA = String(first.dateRaw || '').trim() !== '';
        const hasDateB = String(current.dateRaw || '').trim() !== '';
        const dateMismatch = hasDateA && hasDateB && first.date !== current.date;

        const isTrueDuplicate = grossSame && customerSame;
        if (isTrueDuplicate) {
          const issue = `Row ${current.row}: duplicate value ${value} (first seen at row ${first.row})`;
          issues.push(issue);
          duplicates.push({
            row: current.row,
            value,
            firstRow: first.row,
            issue,
          });
        } else {
          const rollbackIssue = `Row ${current.row}: roll back OR# ${value} (mismatch with row ${first.row})`;
          issues.push(rollbackIssue);
          rollbacks.push({
            row: current.row,
            from: first.receipt,
            to: current.receipt,
            issue: rollbackIssue,
            type: 'duplicate_mismatch',
          });
        }

        if (grossMismatch || customerMismatch || dateMismatch) {
          const parts = [];
          if (grossMismatch) {
            parts.push(`gross (${first.row}=${roundAyalaAmount(first.gross).toFixed(2)}, ${current.row}=${roundAyalaAmount(current.gross).toFixed(2)})`);
          }
          if (customerMismatch) {
            parts.push(`customer (${first.row}="${first.customerRaw}", ${current.row}="${current.customerRaw}")`);
          }
          if (dateMismatch) {
            parts.push(`date (${first.row}="${first.dateRaw}", ${current.row}="${current.dateRaw}")`);
          }
          const detailIssue = `Row ${current.row}: duplicate OR# ${value} mismatch -> ${parts.join(', ')}`;
          issues.push(detailIssue);
          duplicateDetailMismatches.push({
            row: current.row,
            firstRow: first.row,
            value,
            grossMismatch,
            customerMismatch,
            dateMismatch,
            issue: detailIssue,
          });
        }
      }
    }

    for (let i = 1; i < entries.length; i += 1) {
      const previous = entries[i - 1].receipt;
      const current = entries[i].receipt;

      if (current < previous) {
        const issue = `Row ${entries[i].row}: out-of-order (${previous} -> ${current})`;
        issues.push(issue);
        rollbacks.push({
          row: entries[i].row,
          from: previous,
          to: current,
          issue,
        });
      }
    }

    // Detect true missing OR# across the full observed range,
    // even when rows are out-of-order (rollback cases).
    const sortedUnique = Array.from(new Set(entries.map((item) => item.receipt)))
      .filter((value) => Number.isFinite(value))
      .sort((a, b) => a - b);
    for (let i = 1; i < sortedUnique.length; i += 1) {
      const previous = sortedUnique[i - 1];
      const current = sortedUnique[i];
      if (current - previous > 1) {
        for (let n = previous + 1; n < current; n += 1) {
          missing.push(n);
        }
      }
    }

    return { missing, issues, duplicates, rollbacks, duplicateDetailMismatches };
  }

  function detectCsvDelimiter(line) {
    if (line.includes('\t')) {
      return '\t';
    }
    if (line.includes(';')) {
      return ';';
    }
    if (line.includes(',')) {
      return ',';
    }
    return null;
  }

  function normalizeAyalaHeader(value) {
    return String(value || '')
      .toUpperCase()
      .replace(/[^A-Z0-9]+/g, '_')
      .replace(/^_+|_+$/g, '');
  }

  function normalizeAyalaFieldName(value) {
    const normalized = normalizeAyalaHeader(value);
    const aliases = {
      LOCAL_TAX: 'LOCAL_TAX',
      LOCAL: 'LOCAL_TAX',
      OTHERSL_SLS: 'OTHER_SLS',
      OTHER_SLS: 'OTHER_SLS',
      OPEN_SALES: 'OPEN_SALES',
      NO_VATEXEMT: 'NO_VATEXEMT',
    };
    return aliases[normalized] || normalized;
  }

  function parseNumericAyala(value) {
    if (value === null || value === undefined) {
      return 0;
    }
    const text = String(value).replace(/,/g, '').trim();
    if (!text) {
      return 0;
    }
    const parenMatch = text.match(/^\((.+)\)$/);
    const normalized = parenMatch ? `-${parenMatch[1]}` : text;
    const parsed = Number.parseFloat(normalized);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function getAyalaNumericFieldValue(record, field) {
    if (!record) {
      return 0;
    }
    const primary = parseNumericAyala(record[field]);
    if (primary !== 0) {
      return primary;
    }
    if (field === 'OTHERSL_SLS') {
      return parseNumericAyala(record.OTHER_SLS);
    }
    if (field === 'OTHER_SLS') {
      return parseNumericAyala(record.OTHERSL_SLS);
    }
    return primary;
  }

  const AYALA_PRODUCT_LINE_FIELDS = new Set(['QTY', 'ITEMCODE', 'PRICE', 'LDISC']);

  function getAyalaRawFieldValues(record, field) {
    if (!record) {
      return [];
    }
    const key = String(field || '').trim();
    const multiLine = record.__multiLineFields;
    if (multiLine && typeof multiLine === 'object' && Array.isArray(multiLine[key]) && multiLine[key].length > 0) {
      return multiLine[key];
    }
    if (key === 'OTHERSL_SLS') {
      const primary = record.OTHERSL_SLS ?? record.OTHER_SLS ?? '';
      return primary === '' ? [] : [primary];
    }
    if (key === 'OTHER_SLS') {
      const primary = record.OTHER_SLS ?? record.OTHERSL_SLS ?? '';
      return primary === '' ? [] : [primary];
    }
    const primary = record[key];
    return (primary === undefined || primary === null || String(primary) === '') ? [] : [primary];
  }

  function getAyalaRawFieldValue(record, field) {
    if (!record) {
      return '';
    }
    const primary = record[field];
    if (primary !== undefined && primary !== null && String(primary) !== '') {
      return primary;
    }
    if (field === 'OTHERSL_SLS') {
      return record.OTHER_SLS ?? '';
    }
    if (field === 'OTHER_SLS') {
      return record.OTHERSL_SLS ?? '';
    }
    return '';
  }

  function getAyalaPaymentDiscrepancy(payment, paymentTotal) {
    return getAyalaFlexibleDiscrepancy(payment, paymentTotal, { allowOppositeSignSameAbs: true });
  }

  function getAyalaFlexibleDiscrepancy(leftValue, rightValue, options = {}) {
    const {
      allowOppositeSignSameAbs = false,
    } = options;
    const leftRounded = roundAyalaAmount(leftValue);
    const rightRounded = roundAyalaAmount(rightValue);
    let discrepancy = Math.abs(leftRounded - rightRounded);
    if (
      allowOppositeSignSameAbs
      && Math.sign(leftRounded) !== 0
      && Math.sign(rightRounded) !== 0
      && Math.sign(leftRounded) !== Math.sign(rightRounded)
      && isSameAyalaAmount(Math.abs(leftValue), Math.abs(rightValue))
    ) {
      discrepancy = 0;
    }
    return discrepancy;
  }

  function getAyalaAdjustedPaymentFieldValue(record, field) {
    const key = String(field || '').trim().toUpperCase();
    const gcExcess = parseNumericAyala(record && record.GC_EXCESS);
    if (key === 'GC_SLS') {
      return getAyalaNumericFieldValue(record, 'GC_SLS') - gcExcess;
    }
    if (key === 'GC_EXCESS') {
      return 0;
    }
    return getAyalaNumericFieldValue(record, field);
  }

  function parseAyalaTableFromRows(rows) {
    if (!rows || rows.length === 0) {
      return { headers: [], records: [] };
    }

    let headerIndex = -1;
    let normalizedHeaders = [];
    for (let i = 0; i < Math.min(rows.length, 30); i += 1) {
      const row = rows[i] || [];
      const candidateHeaders = row.map((cell) => normalizeAyalaFieldName(cell));
      const hasKeyFields =
        candidateHeaders.includes('TRN_DATE')
        && (candidateHeaders.includes('TER_NO') || candidateHeaders.includes('TRANSACTION_NO'));
      if (hasKeyFields) {
        headerIndex = i;
        normalizedHeaders = candidateHeaders;
        break;
      }
    }

    if (headerIndex < 0) {
      normalizedHeaders = (rows[0] || []).map((cell) => normalizeAyalaFieldName(cell));
      headerIndex = 0;
    }

    const records = [];
    for (let r = headerIndex + 1; r < rows.length; r += 1) {
      const row = rows[r] || [];
      const record = {};
      let hasValue = false;
      normalizedHeaders.forEach((header, index) => {
        if (!header) {
          return;
        }
        const value = row[index] ?? '';
        const text = String(value).trim();
        if (text) {
          hasValue = true;
        }
        record[header] = text;
      });
      if (hasValue) {
        records.push(record);
      }
    }

    return { headers: normalizedHeaders.filter(Boolean), records };
  }

  function parseAyalaTableFromText(textContent) {
    const lines = String(textContent || '')
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean);

    if (lines.length === 0) {
      return { headers: [], records: [] };
    }

    // Support Ayala key-value files where each line is "FIELD,VALUE".
    // Example:
    // CCCODE,84306000000000247
    // TRN_DATE,2025-09-09
    // ...
    const kvRows = [];
    let kvLikeCount = 0;
    lines.forEach((line) => {
      const parts = line.split(',');
      if (parts.length < 2) {
        return;
      }
      const key = normalizeAyalaFieldName(parts[0]);
      const values = parts.slice(1).map((part) => String(part || '').trim());
      if (/^[A-Z][A-Z0-9_ ]*$/.test(String(parts[0]).trim())) {
        kvLikeCount += 1;
      }
      if (key) {
        kvRows.push({ key, values });
      }
    });

    const kvRatio = kvRows.length > 0 ? (kvLikeCount / kvRows.length) : 0;
    if (kvRows.length >= 8 && kvRatio >= 0.8) {
      const maxValueColumns = kvRows.reduce((max, row) => Math.max(max, row.values.length), 0);

      // Support EOD "FIELD,val1,val2,...", where each value column is a terminal record.
      if (maxValueColumns > 1) {
        const records = Array.from({ length: maxValueColumns }, () => ({}));
        kvRows.forEach(({ key, values }) => {
          for (let i = 0; i < maxValueColumns; i += 1) {
            records[i][key] = String(values[i] ?? '').trim();
          }
        });
        const filtered = records.filter((record) => {
          return Object.values(record).some((value) => String(value || '').trim() !== '');
        });
        const allHeaders = new Set();
        filtered.forEach((record) => {
          Object.keys(record).forEach((header) => allHeaders.add(header));
        });
        return {
          headers: Array.from(allHeaders),
          records: filtered,
        };
      }

      // Single-value-per-line key/value stream (hourly-like).
      const globalFields = {};
      const records = [];
      let current = {};
      const appendProductLineValue = (target, key, value) => {
        if (!AYALA_PRODUCT_LINE_FIELDS.has(key)) {
          return;
        }
        if (!target.__multiLineFields || typeof target.__multiLineFields !== 'object') {
          target.__multiLineFields = {};
        }
        if (!Array.isArray(target.__multiLineFields[key])) {
          target.__multiLineFields[key] = [];
        }
        if (
          Object.prototype.hasOwnProperty.call(target, key)
          && String(target[key] ?? '').trim() !== ''
          && target.__multiLineFields[key].length === 0
        ) {
          target.__multiLineFields[key].push(String(target[key]).trim());
        }
        if (value !== '') {
          target.__multiLineFields[key].push(value);
        }
      };
      kvRows.forEach(({ key, values }) => {
        const value = String(values[0] ?? '').trim();
        const isGlobalKey = key === 'CCCODE' || key === 'MERCHANT_NAME' || key === 'TRN_DATE';
        if (isGlobalKey) {
          globalFields[key] = value;
        }
        if (key === 'CDATE' && current.TRANSACTION_NO) {
          records.push(current);
          current = { ...globalFields };
        }
        appendProductLineValue(current, key, value);
        current[key] = value;
      });
      if (Object.keys(current).length > 0) {
        records.push(current);
      }
      const allHeaders = new Set();
      records.forEach((record) => {
        Object.keys(record)
          .filter((header) => !String(header || '').startsWith('__'))
          .forEach((header) => allHeaders.add(header));
      });
      return {
        headers: Array.from(allHeaders),
        records,
      };
    }

    const delimiter = detectCsvDelimiter(lines[0]) || ',';
    const rows = lines.map((line) => line.split(delimiter).map((cell) => cell.trim()));
    return parseAyalaTableFromRows(rows);
  }

  function parseAyalaTableFromWorkbook(arrayBuffer) {
    if (typeof XLSX === 'undefined') {
      return { headers: [], records: [], parseError: 'Excel parser is not loaded.' };
    }

    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
    return parseAyalaTableFromRows(rows);
  }

  function roundAyalaAmount(value) {
    return Math.round(value * 100) / 100;
  }

  const AYALA_DISCREPANCY_TOLERANCE = 1.0;
  // Temporary safety switch: keep raw hourly transactions unchanged (no move-pairs).
  const AYALA_MOVE_PAIRS_ENABLED = true;

  function isSameAyalaAmount(a, b) {
    return Math.abs(roundAyalaAmount(a) - roundAyalaAmount(b)) <= AYALA_DISCREPANCY_TOLERANCE;
  }

  function formatAyalaAmount(value) {
    return roundAyalaAmount(value).toFixed(2);
  }

  function getAyalaKey(record) {
    return [
      record.CCCODE || 'UNKNOWN_CCCODE',
      record.TER_NO || 'UNKNOWN_TER',
      record.TRN_DATE || 'UNKNOWN_DATE',
    ].join('|');
  }

  function getAyalaTransactionNo(record) {
    return String(record.TRANSACTION_NO || '').trim();
  }

  function getAyalaTerminalNo(record) {
    return String(record.TER_NO || '').trim();
  }

  function getAyalaTransactionCompositeKey(record) {
    return `${getAyalaTerminalNo(record)}|${getAyalaTransactionNo(record)}`;
  }

  function buildAyalaCombinedFieldValues(records) {
    const result = new Map();
    const source = Array.isArray(records) ? records : [];
    if (source.length === 0) {
      return result;
    }

    const textOnlyFields = new Set([
      'CCCODE', 'MERCHANT_NAME', 'TRN_DATE', 'CDATE', 'TRN_TIME', 'TER_NO',
      'MOBILE_NO', 'TRN_TYPE', 'SLS_FLAG', 'ITEMCODE',
    ]);

    const allFields = [];
    const seenFields = new Set();
    source.forEach((record) => {
      Object.keys(record || {}).forEach((field) => {
        if (!seenFields.has(field)) {
          seenFields.add(field);
          allFields.push(field);
        }
      });
    });

    allFields.forEach((field) => {
      const values = source
        .map((record) => String((record && Object.prototype.hasOwnProperty.call(record, field)) ? record[field] : '').trim())
        .filter((value) => value !== '');

      if (values.length === 0) {
        result.set(field, '');
        return;
      }

      if (field === 'TER_NO') {
        result.set(field, Array.from(new Set(values)).join(', '));
        return;
      }

      if (textOnlyFields.has(field)) {
        result.set(field, Array.from(new Set(values)).join(', '));
        return;
      }

      const numericPattern = /^-?\d+(?:\.\d+)?$/;
      const normalizedValues = values.map((value) => value.replace(/,/g, ''));
      const isFullyNumeric = normalizedValues.every((value) => numericPattern.test(value));
      if (isFullyNumeric) {
        const sum = normalizedValues.reduce((acc, value) => acc + (Number.parseFloat(value) || 0), 0);
        const hasDecimal = normalizedValues.some((value) => value.includes('.'));
        result.set(field, hasDecimal ? formatAyalaAmount(sum) : String(Math.round(sum)));
        return;
      }

      result.set(field, Array.from(new Set(values)).join(', '));
    });

    return result;
  }

  function compareAyalaTxnEntries(aTer, aTxn, bTer, bTxn) {
    const aTerNum = Number.parseInt(aTer, 10);
    const bTerNum = Number.parseInt(bTer, 10);
    if (Number.isFinite(aTerNum) && Number.isFinite(bTerNum) && aTerNum !== bTerNum) {
      return aTerNum - bTerNum;
    }
    if (aTer !== bTer) {
      return aTer.localeCompare(bTer);
    }
    const aTxnNum = Number.parseInt(aTxn, 10);
    const bTxnNum = Number.parseInt(bTxn, 10);
    if (Number.isFinite(aTxnNum) && Number.isFinite(bTxnNum)) {
      return aTxnNum - bTxnNum;
    }
    return aTxn.localeCompare(bTxn);
  }

  function buildAyalaRefundPairPlan(hourlyRecords) {
    if (!AYALA_MOVE_PAIRS_ENABLED) {
      return {
        refundEntries: [],
        refundMatches: [],
        unmatchedRefunds: [],
        movedTxnKeys: new Set(),
        pairSource: [],
      };
    }

    const entries = (hourlyRecords || [])
      .map((record) => {
        const txn = String(record.TRANSACTION_NO || '').trim();
        const ter = String(record.TER_NO || '').trim();
        return {
          key: getAyalaTransactionCompositeKey(record),
          txn,
          ter,
          txnNum: Number.parseInt(txn, 10),
          record,
        };
      })
      .filter((entry) => entry.key && entry.key !== '|');

    const isRefundRecord = (entry) => {
      const gross = parseNumericAyala(entry.record.GROSS_SLS);
      const refundAmt = parseNumericAyala(entry.record.REFUND_AMT);
      const slsFlag = String(entry.record.SLS_FLAG || '').trim().toUpperCase();
      const trnType = String(entry.record.TRN_TYPE || '').trim().toUpperCase();
      const negativeSignalFields = [
        'GROSS_SLS', 'VAT_AMNT', 'VATABLE_SLS', 'NONVAT_SLS', 'VATEXEMPT_SLS', 'VATEXEMPT_AMNT',
        'CASH_SLS', 'CARD_SLS', 'EPAY_SLS', 'DCARD_SLS', 'OTHER_SLS', 'CHECK_SLS', 'GC_SLS',
        'PAYMAYA_SLS', 'GRAB_SLS', 'FOODPANDA_SLS', 'SNRCIT_DISC', 'PWD_DISC', 'OTHER_DISC',
      ];
      const hasNegativeSignal = negativeSignalFields.some((field) => parseNumericAyala(entry.record[field]) < 0);
      return (
        gross < 0
        || refundAmt > 0
        || slsFlag === 'R'
        || slsFlag === 'REFUND'
        || trnType === 'R'
        || trnType === 'REFUND'
        || hasNegativeSignal
      );
    };

    const refundEntries = entries.filter((entry) => isRefundRecord(entry));
    const originalEntries = entries.filter((entry) => !isRefundRecord(entry));
    const usedOriginalKeys = new Set();
    const refundMatches = [];
    const unmatchedRefunds = [];

    refundEntries.forEach((refund) => {
      const refundItem = String(refund.record.ITEMCODE || '').trim().toUpperCase();
      const refundPrice = Math.abs(parseNumericAyala(refund.record.PRICE));
      const refundQty = Math.abs(parseNumericAyala(refund.record.QTY));
      const refundGross = Math.abs(parseNumericAyala(refund.record.GROSS_SLS));
      const refundVatable = Math.abs(parseNumericAyala(refund.record.VATABLE_SLS));
      const refundVat = Math.abs(parseNumericAyala(refund.record.VAT_AMNT));

      const scoredCandidates = originalEntries
        .filter((orig) => {
          if (usedOriginalKeys.has(orig.key)) {
            return false;
          }
          if ((orig.ter || '-') !== (refund.ter || '-')) {
            return false;
          }
          if (Number.isFinite(orig.txnNum) && Number.isFinite(refund.txnNum)) {
            return orig.txnNum <= refund.txnNum;
          }
          return true;
        })
        .map((orig) => {
          const origItem = String(orig.record.ITEMCODE || '').trim().toUpperCase();
          const origPrice = Math.abs(parseNumericAyala(orig.record.PRICE));
          const origQty = Math.abs(parseNumericAyala(orig.record.QTY));
          const origGross = Math.abs(parseNumericAyala(orig.record.GROSS_SLS));
          const origVatable = Math.abs(parseNumericAyala(orig.record.VATABLE_SLS));
          const origVat = Math.abs(parseNumericAyala(orig.record.VAT_AMNT));
          let score = 0;

          if (refundItem && origItem && refundItem === origItem) score += 100;
          if (refundPrice > 0 && isSameAyalaAmount(refundPrice, origPrice)) score += 60;
          if (refundQty > 0 && isSameAyalaAmount(refundQty, origQty)) score += 20;
          if (refundGross > 0 && isSameAyalaAmount(refundGross, origGross)) score += 40;
          if (refundVatable > 0 && isSameAyalaAmount(refundVatable, origVatable)) score += 35;
          if (refundVat > 0 && isSameAyalaAmount(refundVat, origVat)) score += 25;

          const txnGap = Number.isFinite(orig.txnNum) && Number.isFinite(refund.txnNum)
            ? Math.abs(refund.txnNum - orig.txnNum)
            : 999999;
          return { orig, score, txnGap };
        })
        .sort((a, b) => {
          if (b.score !== a.score) return b.score - a.score;
          return a.txnGap - b.txnGap;
        });

      const best = scoredCandidates[0];
      if (!best || best.score < 20) {
        unmatchedRefunds.push(refund);
        return;
      }
      usedOriginalKeys.add(best.orig.key);
      refundMatches.push({ refund, original: best.orig });
    });

    const movedTxnKeys = new Set();
    refundMatches.forEach((pair) => {
      movedTxnKeys.add(pair.refund.key);
      movedTxnKeys.add(pair.original.key);
    });

    const pairSource = refundMatches.length > 0
      ? refundMatches
      : unmatchedRefunds.map((refund) => ({ refund, original: null }));

    return {
      refundEntries,
      refundMatches,
      unmatchedRefunds,
      movedTxnKeys,
      pairSource,
    };
  }

  function buildAyalaHourlySummary(hourlyRecords) {
    const groups = new Map();
    const numericFields = new Set([
      'GROSS_SLS',
      'VAT_AMNT',
      'VATABLE_SLS',
      'NONVAT_SLS',
      'VATEXEMPT_SLS',
      'VATEXEMPT_AMNT',
      'LOCAL_TAX',
      'PWD_DISC',
      'SNRCIT_DISC',
      'EMPLO_DISC',
      'AYALA_DISC',
      'STORE_DISC',
      'OTHER_DISC',
      'REFUND_AMT',
      'SCHRGE_AMT',
      'OTHER_SCHR',
      'CASH_SLS',
      'CARD_SLS',
      'EPAY_SLS',
      'DCARD_SLS',
      'OTHER_SLS',
      'CHECK_SLS',
      'GC_SLS',
      'MASTERCARD_SLS',
      'VISA_SLS',
      'AMEX_SLS',
      'DINERS_SLS',
      'JCB_SLS',
      'GCASH_SLS',
      'PAYMAYA_SLS',
      'ALIPAY_SLS',
      'WECHAT_SLS',
      'GRAB_SLS',
      'FOODPANDA_SLS',
      'MASTERDEBIT_SLS',
      'VISADEBIT_SLS',
      'PAYPAL_SLS',
      'ONLINE_SLS',
      'OPEN_SALES',
      'OPEN_SALES_2',
      'OPEN_SALES_3',
      'OPEN_SALES_4',
      'OPEN_SALES_5',
      'OPEN_SALES_6',
      'OPEN_SALES_7',
      'OPEN_SALES_8',
      'OPEN_SALES_9',
      'OPEN_SALES_10',
      'OPEN_SALES_11',
      'GC_EXCESS',
    ]);

    hourlyRecords.forEach((record) => {
      const key = getAyalaKey(record);
      if (!groups.has(key)) {
        groups.set(key, {
          key,
          sample: record,
          transactions: new Set(),
          minTxn: null,
          maxTxn: null,
          sums: {},
          rowCount: 0,
        });
      }

      const group = groups.get(key);
      group.rowCount += 1;
      const txn = getAyalaTransactionNo(record);
      if (txn) {
        group.transactions.add(txn);
        if (/^\d+$/.test(txn)) {
          const n = Number.parseInt(txn, 10);
          if (group.minTxn === null || n < group.minTxn) {
            group.minTxn = n;
          }
          if (group.maxTxn === null || n > group.maxTxn) {
            group.maxTxn = n;
          }
        }
      }

      numericFields.forEach((field) => {
        const numericValue = field === 'OTHER_SLS'
          ? getAyalaNumericFieldValue(record, 'OTHER_SLS')
          : parseNumericAyala(record[field]);
        group.sums[field] = (group.sums[field] || 0) + numericValue;
      });
    });

    return groups;
  }

  function buildAyalaEodMap(eodRecords) {
    const map = new Map();
    eodRecords.forEach((record) => {
      map.set(getAyalaKey(record), record);
    });
    return map;
  }

  function validateAyalaData(hourlyRecords, eodRecords) {
    const isZeroLikeTxnRef = (value) => {
      const text = String(value || '').trim();
      if (!text) {
        return true;
      }
      const parsed = Number.parseInt(text.replace(/,/g, ''), 10);
      return Number.isFinite(parsed) && parsed === 0;
    };
    const shouldIgnoreMissingHourlyGroup = (eodRecord) => {
      if (!eodRecord || typeof eodRecord !== 'object') {
        return false;
      }
      const noTrn = Number.parseInt(String(eodRecord.NO_TRN || '').replace(/,/g, ''), 10);
      const gross = parseNumericAyala(eodRecord.GROSS_SLS);
      const epay = parseNumericAyala(eodRecord.EPAY_SLS);
      const refund = parseNumericAyala(eodRecord.REFUND_AMT);
      return Number.isFinite(noTrn)
        && noTrn === 0
        && Math.abs(gross) <= AYALA_DISCREPANCY_TOLERANCE
        && Math.abs(epay) <= AYALA_DISCREPANCY_TOLERANCE
        && Math.abs(refund) <= AYALA_DISCREPANCY_TOLERANCE
        && isZeroLikeTxnRef(eodRecord.STRANS)
        && isZeroLikeTxnRef(eodRecord.ETRANS);
    };
    const duplicateSensitiveFields = [
      'GROSS_SLS', 'VAT_AMNT', 'VATABLE_SLS', 'NONVAT_SLS', 'VATEXEMPT_SLS', 'VATEXEMPT_AMNT',
      'LOCAL_TAX', 'PWD_DISC', 'SNRCIT_DISC', 'EMPLO_DISC', 'AYALA_DISC', 'STORE_DISC', 'OTHER_DISC',
      'REFUND_AMT', 'SCHRGE_AMT', 'OTHER_SCHR', 'CASH_SLS', 'CARD_SLS', 'EPAY_SLS', 'DCARD_SLS',
      'OTHER_SLS', 'OTHERSL_SLS', 'CHECK_SLS', 'GC_SLS', 'MASTERCARD_SLS', 'VISA_SLS', 'AMEX_SLS',
      'DINERS_SLS', 'JCB_SLS', 'GCASH_SLS', 'PAYMAYA_SLS', 'ALIPAY_SLS', 'WECHAT_SLS', 'GRAB_SLS',
      'FOODPANDA_SLS', 'MASTERDEBIT_SLS', 'VISADEBIT_SLS', 'PAYPAL_SLS', 'ONLINE_SLS',
      'OPEN_SALES', 'OPEN_SALES_2', 'OPEN_SALES_3', 'OPEN_SALES_4', 'OPEN_SALES_5', 'OPEN_SALES_6',
      'OPEN_SALES_7', 'OPEN_SALES_8', 'OPEN_SALES_9', 'OPEN_SALES_10', 'OPEN_SALES_11', 'GC_EXCESS',
      'TRN_TYPE', 'SLS_FLAG',
    ];
    const normalizeForDuplicate = (record, field) => {
      if (field === 'OTHER_SLS' || field === 'OTHERSL_SLS') {
        return roundAyalaAmount(getAyalaNumericFieldValue(record, field)).toFixed(2);
      }
      const rawValue = getAyalaRawFieldValue(record, field);
      const numeric = parseNumericAyala(rawValue);
      if (String(rawValue || '').trim() !== '' && Number.isFinite(numeric) && /^-?\d+(?:\.\d+)?$/.test(String(rawValue).trim())) {
        return roundAyalaAmount(numeric).toFixed(2);
      }
      return String(rawValue || '').trim().toUpperCase();
    };
    const buildDuplicateSignature = (record) => {
      return duplicateSensitiveFields
        .map((field) => `${field}:${normalizeForDuplicate(record, field)}`)
        .join('|');
    };
    const byTxnComposite = new Map();
    (hourlyRecords || []).forEach((record) => {
      const key = getAyalaTransactionCompositeKey(record);
      if (!key || key === '|') {
        return;
      }
      if (!byTxnComposite.has(key)) {
        byTxnComposite.set(key, []);
      }
      byTxnComposite.get(key).push(record);
    });
    const validationHourlyRecords = [];
    byTxnComposite.forEach((records) => {
      if (Array.isArray(records) && records.length > 0) {
        validationHourlyRecords.push(records[0]);
      }
    });

    const hourlyGroups = buildAyalaHourlySummary(validationHourlyRecords);
    const refundPlan = buildAyalaRefundPairPlan(validationHourlyRecords);
    // Keep validation totals aligned with checker workbook when refund pairs exist.
    const APPLY_MOVE_PAIR_ADJUSTMENTS_IN_VALIDATION = true;
    const movedTxnKeysForValidation = APPLY_MOVE_PAIR_ADJUSTMENTS_IN_VALIDATION
      ? new Set(
        (refundPlan.refundMatches || []).flatMap((pair) => ([
          String(pair && pair.refund ? pair.refund.key : '').trim(),
          String(pair && pair.original ? pair.original.key : '').trim(),
        ])).filter((key) => key && key !== '|')
      )
      : new Set();
    const adjustmentFields = [
      'GROSS_SLS', 'VAT_AMNT', 'VATABLE_SLS', 'NONVAT_SLS', 'VATEXEMPT_SLS', 'VATEXEMPT_AMNT',
      'LOCAL_TAX', 'PWD_DISC', 'SNRCIT_DISC', 'EMPLO_DISC', 'AYALA_DISC', 'STORE_DISC',
      'OTHER_DISC', 'REFUND_AMT', 'SCHRGE_AMT', 'OTHER_SCHR',
    ];
    const paymentAdjustmentFields = [
      'CASH_SLS', 'OTHER_SLS', 'CHECK_SLS', 'GC_SLS', 'MASTERCARD_SLS', 'VISA_SLS',
      'AMEX_SLS', 'DINERS_SLS', 'JCB_SLS', 'GCASH_SLS', 'PAYMAYA_SLS', 'ALIPAY_SLS',
      'WECHAT_SLS', 'GRAB_SLS', 'FOODPANDA_SLS', 'MASTERDEBIT_SLS', 'VISADEBIT_SLS',
      'PAYPAL_SLS', 'ONLINE_SLS', 'OPEN_SALES', 'OPEN_SALES_2', 'OPEN_SALES_3',
      'OPEN_SALES_4', 'OPEN_SALES_5', 'OPEN_SALES_6', 'OPEN_SALES_7', 'OPEN_SALES_8',
      'OPEN_SALES_9', 'OPEN_SALES_10', 'OPEN_SALES_11', 'GC_EXCESS',
    ];
    const movedOriginalGrossTotal = APPLY_MOVE_PAIR_ADJUSTMENTS_IN_VALIDATION
      ? roundAyalaAmount(
        (refundPlan.refundMatches || []).reduce(
          (acc, pair) => acc + parseNumericAyala(pair && pair.original ? pair.original.record.GROSS_SLS : 0),
          0
        )
      )
      : 0;
    const movedRefundTotal = APPLY_MOVE_PAIR_ADJUSTMENTS_IN_VALIDATION
      ? roundAyalaAmount(
        (refundPlan.refundMatches || []).reduce(
          (acc, pair) => acc + Math.abs(parseNumericAyala(pair && pair.refund ? pair.refund.record.REFUND_AMT : 0)),
          0
        )
      )
      : 0;
    const hourlyAdjustByField = {};
    adjustmentFields.forEach((field) => {
      hourlyAdjustByField[field] = 0;
    });
    const paymentAdjustByField = {};
    paymentAdjustmentFields.forEach((field) => {
      paymentAdjustByField[field] = 0;
    });
    const terminalGrossAdjustments = new Map();

    if (APPLY_MOVE_PAIR_ADJUSTMENTS_IN_VALIDATION) {
      refundPlan.pairSource.forEach((pair) => {
        const refund = pair.refund;
        const original = pair.original;
        adjustmentFields.forEach((field) => {
          const delta = parseNumericAyala(refund.record[field]) - parseNumericAyala(original ? original.record[field] : 0);
          hourlyAdjustByField[field] += delta;
        });
        paymentAdjustmentFields.forEach((field) => {
          const delta = getAyalaAdjustedPaymentFieldValue(refund.record, field)
            + getAyalaAdjustedPaymentFieldValue(original ? original.record : null, field);
          paymentAdjustByField[field] += delta;
        });
        const terKey = String(refund.ter || '-');
        terminalGrossAdjustments.set(
          terKey,
          (terminalGrossAdjustments.get(terKey) || 0) + parseNumericAyala(original ? original.record.GROSS_SLS : 0)
        );
      });
    }

    // Keep grouped field comparison based on full/raw hourly sums.
    // This avoids false field mismatches (e.g. GROSS_SLS) when moved-pair adjustment
    // is intended only for checker-side matrix layout and summary reconciliation.

    const eodMap = buildAyalaEodMap(eodRecords);
    const mismatches = [];
    const missingInEod = [];
    const missingInHourly = [];
    const comparedFields = [
      'GROSS_SLS',
      'VAT_AMNT',
      'VATABLE_SLS',
      'NONVAT_SLS',
      'VATEXEMPT_SLS',
      'VATEXEMPT_AMNT',
      'LOCAL_TAX',
      'PWD_DISC',
      'SNRCIT_DISC',
      'EMPLO_DISC',
      'AYALA_DISC',
      'STORE_DISC',
      'OTHER_DISC',
      'REFUND_AMT',
      'SCHRGE_AMT',
      'OTHER_SCHR',
      'CASH_SLS',
      'CARD_SLS',
      'EPAY_SLS',
      'DCARD_SLS',
      'OTHER_SLS',
      'CHECK_SLS',
      'GC_SLS',
      'MASTERCARD_SLS',
      'VISA_SLS',
      'AMEX_SLS',
      'DINERS_SLS',
      'JCB_SLS',
      'GCASH_SLS',
      'PAYMAYA_SLS',
      'ALIPAY_SLS',
      'WECHAT_SLS',
      'GRAB_SLS',
      'FOODPANDA_SLS',
      'MASTERDEBIT_SLS',
      'VISADEBIT_SLS',
      'PAYPAL_SLS',
      'ONLINE_SLS',
      'OPEN_SALES',
      'OPEN_SALES_2',
      'OPEN_SALES_3',
      'OPEN_SALES_4',
      'OPEN_SALES_5',
      'OPEN_SALES_6',
      'OPEN_SALES_7',
      'OPEN_SALES_8',
      'OPEN_SALES_9',
      'OPEN_SALES_10',
      'OPEN_SALES_11',
      'GC_EXCESS',
    ];
    const totals = {
      grossHourly: 0,
      grossEod: 0,
      refundAmountHourly: 0,
      refundAmountEod: 0,
      refundTxnHourly: 0,
      refundTxnEod: 0,
      hourlyTxnCount: 0,
      epayHourly: 0,
      epayEod: 0,
      salesTotalDaily: 0,
      salesTotalTransaction: 0,
      oldGrnTotEod: 0,
      newGrnTotEod: 0,
      vatAmntHourly: 0,
      vatableSlsHourly: 0,
      nonvatSlsHourly: 0,
      vatexemptSlsHourly: 0,
    };
    const paymentFields = [
      'CASH_SLS', 'OTHER_SLS', 'CHECK_SLS', 'GC_SLS', 'MASTERCARD_SLS', 'VISA_SLS',
      'AMEX_SLS', 'DINERS_SLS', 'JCB_SLS', 'GCASH_SLS', 'PAYMAYA_SLS', 'ALIPAY_SLS',
      'WECHAT_SLS', 'GRAB_SLS', 'FOODPANDA_SLS', 'MASTERDEBIT_SLS', 'VISADEBIT_SLS',
      'PAYPAL_SLS', 'ONLINE_SLS', 'OPEN_SALES', 'OPEN_SALES_2', 'OPEN_SALES_3',
      'OPEN_SALES_4', 'OPEN_SALES_5', 'OPEN_SALES_6', 'OPEN_SALES_7', 'OPEN_SALES_8',
      'OPEN_SALES_9', 'OPEN_SALES_10', 'OPEN_SALES_11', 'GC_EXCESS',
    ];
    const paymentDailyBreakdown = {};
    const paymentTransactionBreakdown = {};
    paymentFields.forEach((field) => {
      paymentDailyBreakdown[field] = 0;
      paymentTransactionBreakdown[field] = 0;
    });
    const hourlyRows = validationHourlyRecords.map((record) => ({
      trnTime: String(record.TRN_TIME || '').trim(),
      transactionNo: String(record.TRANSACTION_NO || '').trim(),
      noTrn: String(record.NO_TRN || '').trim(),
      terNo: String(record.TER_NO || '').trim(),
      gross: roundAyalaAmount(parseNumericAyala(record.GROSS_SLS)),
      refund: roundAyalaAmount(parseNumericAyala(record.REFUND_AMT)),
    }));

    const duplicateTxnMap = new Map();
    (hourlyRecords || []).forEach((record) => {
      const terNo = String(record.TER_NO || '').trim() || '-';
      const txnNo = String(record.TRANSACTION_NO || '').trim();
      if (!txnNo) {
        return;
      }
      const dupKey = `${terNo}|${txnNo}`;
      if (!duplicateTxnMap.has(dupKey)) {
        duplicateTxnMap.set(dupKey, {
          terNo,
          txnNo,
          count: 0,
        });
      }
      duplicateTxnMap.get(dupKey).count += 1;
    });
    const duplicateTransactions = Array.from(duplicateTxnMap.entries())
      .map(([dupKey, item]) => {
        const records = byTxnComposite.get(dupKey) || [];
        const signatures = new Set(records.map((record) => buildDuplicateSignature(record)));
        const hasRealConflict = signatures.size > 1;
        return {
          ...item,
          hasRealConflict,
        };
      })
      .filter((item) => item.count > 1 && item.hasRealConflict)
      .sort((a, b) => compareAyalaTxnEntries(a.terNo, a.txnNo, b.terNo, b.txnNo));

    hourlyRows.sort((a, b) => {
      const aTxn = Number.parseInt(a.transactionNo, 10);
      const bTxn = Number.parseInt(b.transactionNo, 10);
      if (Number.isFinite(aTxn) && Number.isFinite(bTxn)) {
        return aTxn - bTxn;
      }
      return a.transactionNo.localeCompare(b.transactionNo);
    });

    // Refund transaction count should follow full hourly transaction set (same as checker NO_TRN basis).
    validationHourlyRecords.forEach((record) => {
      const refundAmt = parseNumericAyala(record.REFUND_AMT);
      const slsFlag = String(record.SLS_FLAG || '').trim().toUpperCase();
      const hasRefundFlag = slsFlag === 'R' || slsFlag === 'REFUND';
      if (hasRefundFlag || refundAmt > 0) {
        totals.refundTxnHourly += 1;
      }
    });

    let movedRowsExcludedForTotals = 0;
    validationHourlyRecords.forEach((record) => {
      const txnComposite = getAyalaTransactionCompositeKey(record);
      if (movedTxnKeysForValidation.has(txnComposite)) {
        movedRowsExcludedForTotals += 1;
        return;
      }
      const refundAmt = parseNumericAyala(record.REFUND_AMT);
      totals.refundAmountHourly += refundAmt;
      totals.grossHourly += parseNumericAyala(record.GROSS_SLS);
      totals.vatAmntHourly += parseNumericAyala(record.VAT_AMNT);
      totals.vatableSlsHourly += parseNumericAyala(record.VATABLE_SLS);
      totals.nonvatSlsHourly += parseNumericAyala(record.NONVAT_SLS);
      totals.vatexemptSlsHourly += parseNumericAyala(record.VATEXEMPT_SLS);
      totals.salesTotalTransaction += parseNumericAyala(record.VAT_AMNT)
        + parseNumericAyala(record.VATABLE_SLS)
        + parseNumericAyala(record.NONVAT_SLS)
        + parseNumericAyala(record.VATEXEMPT_SLS);
      totals.epayHourly += parseNumericAyala(record.EPAY_SLS);
      paymentFields.forEach((field) => {
        paymentTransactionBreakdown[field] += getAyalaAdjustedPaymentFieldValue(record, field);
      });
    });
    if (APPLY_MOVE_PAIR_ADJUSTMENTS_IN_VALIDATION && movedRowsExcludedForTotals > 0) {
      // Keep aligned with checker formula behavior:
      // GROSS adds moved original gross once; REFUND adds absolute refund total once.
      totals.grossHourly += movedOriginalGrossTotal;
      totals.refundAmountHourly += movedRefundTotal;
      totals.vatAmntHourly += hourlyAdjustByField.VAT_AMNT || 0;
      totals.vatableSlsHourly += hourlyAdjustByField.VATABLE_SLS || 0;
      totals.nonvatSlsHourly += hourlyAdjustByField.NONVAT_SLS || 0;
      totals.vatexemptSlsHourly += hourlyAdjustByField.VATEXEMPT_SLS || 0;
      totals.salesTotalTransaction += (hourlyAdjustByField.VAT_AMNT || 0)
        + (hourlyAdjustByField.VATABLE_SLS || 0)
        + (hourlyAdjustByField.NONVAT_SLS || 0)
        + (hourlyAdjustByField.VATEXEMPT_SLS || 0);
      paymentFields.forEach((field) => {
        paymentTransactionBreakdown[field] += paymentAdjustByField[field] || 0;
      });
      paymentTransactionBreakdown.OTHER_SCHR = (paymentTransactionBreakdown.OTHER_SCHR || 0) + (hourlyAdjustByField.OTHER_SCHR || 0);
    }

    totals.hourlyTxnCount = new Set(
      validationHourlyRecords
        .map((record) => getAyalaTransactionCompositeKey(record))
        .filter((value) => value && value !== '|')
    ).size || validationHourlyRecords.length;

    const terminalMap = new Map();
    const ensureTerminal = (terNo) => {
      const key = String(terNo || '').trim() || '-';
      if (!terminalMap.has(key)) {
        terminalMap.set(key, {
          terNo: key,
          hourlyGross: 0,
          eodGross: 0,
          hourlyEpay: 0,
          eodEpay: 0,
          hourlyTxnSet: new Set(),
          eodNoTrn: 0,
        });
      }
      return terminalMap.get(key);
    };

    eodRecords.forEach((record) => {
      totals.grossEod += parseNumericAyala(record.GROSS_SLS);
      totals.refundAmountEod += parseNumericAyala(record.REFUND_AMT);
      totals.refundTxnEod += Number.parseInt(String(record.NO_REFUND || '').replace(/,/g, ''), 10) || 0;
      totals.epayEod += parseNumericAyala(record.EPAY_SLS);
      totals.salesTotalDaily += parseNumericAyala(record.NEW_GRNTOT) - parseNumericAyala(record.OLD_GRNTOT);
      totals.oldGrnTotEod += parseNumericAyala(record.OLD_GRNTOT);
      totals.newGrnTotEod += parseNumericAyala(record.NEW_GRNTOT);
      paymentFields.forEach((field) => {
        paymentDailyBreakdown[field] += getAyalaAdjustedPaymentFieldValue(record, field);
      });

      const terminal = ensureTerminal(record.TER_NO);
      terminal.eodGross += parseNumericAyala(record.GROSS_SLS);
      terminal.eodEpay += parseNumericAyala(record.EPAY_SLS);
      terminal.eodNoTrn += Number.parseInt(String(record.NO_TRN || '').replace(/,/g, ''), 10) || 0;
    });

    validationHourlyRecords.forEach((record) => {
      const txnComposite = getAyalaTransactionCompositeKey(record);
      const terminal = ensureTerminal(record.TER_NO);
      if (!movedTxnKeysForValidation.has(txnComposite)) {
        terminal.hourlyGross += parseNumericAyala(record.GROSS_SLS);
      }
      terminal.hourlyEpay += parseNumericAyala(record.EPAY_SLS);
      if (txnComposite && txnComposite !== '|') {
        terminal.hourlyTxnSet.add(txnComposite);
      }
    });
    if (APPLY_MOVE_PAIR_ADJUSTMENTS_IN_VALIDATION && movedRowsExcludedForTotals > 0) {
      terminalGrossAdjustments.forEach((grossAdjust, terNo) => {
        const terminal = ensureTerminal(terNo);
        terminal.hourlyGross += grossAdjust;
      });
    }

    hourlyGroups.forEach((group, key) => {
      const eod = eodMap.get(key);
      if (!eod) {
        missingInEod.push(key);
        return;
      }

      const issues = [];
      const hourlyTrnCount = group.transactions.size > 0 ? group.transactions.size : group.rowCount;
      const eodNoTrn = Number.parseInt(String(eod.NO_TRN || '').replace(/,/g, ''), 10);
      if (Number.isFinite(eodNoTrn) && eodNoTrn !== hourlyTrnCount) {
        issues.push(`NO_TRN mismatch: hourly=${hourlyTrnCount}, eod=${eodNoTrn}`);
      }

      const eodStart = Number.parseInt(String(eod.STRANS || '').replace(/,/g, ''), 10);
      const eodEnd = Number.parseInt(String(eod.ETRANS || '').replace(/,/g, ''), 10);
      if (Number.isFinite(eodStart) && group.minTxn !== null && eodStart !== group.minTxn) {
        issues.push(`STRANS mismatch: hourly=${group.minTxn}, eod=${eodStart}`);
      }
      if (Number.isFinite(eodEnd) && group.maxTxn !== null && eodEnd !== group.maxTxn) {
        issues.push(`ETRANS mismatch: hourly=${group.maxTxn}, eod=${eodEnd}`);
      }

      comparedFields.forEach((field) => {
        const hourlyValue = roundAyalaAmount(group.sums[field] || 0);
        const eodValue = roundAyalaAmount(parseNumericAyala(eod[field]));
        const shouldCompare = Number.isFinite(eodValue) && (eod[field] !== undefined && String(eod[field]).trim() !== '');
        const fieldDiscrepancy = getAyalaFlexibleDiscrepancy(hourlyValue, eodValue, { allowOppositeSignSameAbs: true });
        if (shouldCompare && fieldDiscrepancy > AYALA_DISCREPANCY_TOLERANCE) {
          issues.push(`${field} mismatch: hourly=${hourlyValue.toFixed(2)}, eod=${eodValue.toFixed(2)}`);
        }
      });

      if (issues.length > 0) {
        mismatches.push({ key, issues });
      }
    });

    eodMap.forEach((eodRecord, key) => {
      if (!hourlyGroups.has(key)) {
        if (shouldIgnoreMissingHourlyGroup(eodRecord)) {
          return;
        }
        missingInHourly.push(key);
      }
    });

    return {
      missingInEod,
      missingInHourly,
      mismatches,
      hourlyGroupCount: hourlyGroups.size,
      eodGroupCount: eodMap.size,
      hourlyRows,
      eodSummary: {
        noTrn: eodRecords.reduce((sum, record) => (
          sum + (Number.parseInt(String(record.NO_TRN || '').replace(/,/g, ''), 10) || 0)
        ), 0),
        strans: eodRecords.map((record) => String(record.STRANS || '').trim()).filter(Boolean),
        etrans: eodRecords.map((record) => String(record.ETRANS || '').trim()).filter(Boolean),
        trnDate: eodRecords.map((record) => String(record.TRN_DATE || '').trim()).find(Boolean) || '',
        trnDates: Array.from(new Set(eodRecords.map((record) => String(record.TRN_DATE || '').trim()).filter(Boolean))),
        terNo: eodRecords.map((record) => String(record.TER_NO || '').trim()).find(Boolean) || '',
        terNos: Array.from(new Set(eodRecords.map((record) => String(record.TER_NO || '').trim()).filter(Boolean))),
      },
      totals: {
        grossHourly: roundAyalaAmount(totals.grossHourly),
        grossEod: roundAyalaAmount(totals.grossEod),
        refundAmountHourly: roundAyalaAmount(totals.refundAmountHourly),
        refundAmountEod: roundAyalaAmount(totals.refundAmountEod),
        refundTxnHourly: totals.refundTxnHourly,
        refundTxnEod: totals.refundTxnEod,
        hourlyTxnCount: totals.hourlyTxnCount,
        epayHourly: roundAyalaAmount(totals.epayHourly),
        epayEod: roundAyalaAmount(totals.epayEod),
        salesTotalDaily: roundAyalaAmount(totals.salesTotalDaily),
        salesTotalTransaction: roundAyalaAmount(totals.salesTotalTransaction),
        oldGrnTotEod: roundAyalaAmount(totals.oldGrnTotEod),
        newGrnTotEod: roundAyalaAmount(totals.newGrnTotEod),
        vatAmntHourly: roundAyalaAmount(totals.vatAmntHourly),
        vatableSlsHourly: roundAyalaAmount(totals.vatableSlsHourly),
        nonvatSlsHourly: roundAyalaAmount(totals.nonvatSlsHourly),
        vatexemptSlsHourly: roundAyalaAmount(totals.vatexemptSlsHourly),
      },
      paymentDailyBreakdown,
      paymentTransactionBreakdown,
      duplicateTransactions,
      terminalSummaries: Array.from(terminalMap.values()).map((item) => ({
        terNo: item.terNo,
        hourlyGross: roundAyalaAmount(item.hourlyGross),
        eodGross: roundAyalaAmount(item.eodGross),
        hourlyEpay: roundAyalaAmount(item.hourlyEpay),
        eodEpay: roundAyalaAmount(item.eodEpay),
        hourlyTxn: item.hourlyTxnSet.size,
        eodNoTrn: item.eodNoTrn,
        duplicateTxnCount: duplicateTransactions.filter((dup) => dup.terNo === item.terNo).length,
      })),
    };
  }

  function buildAyalaValidationReport(result) {
    const lines = [];
    const reportedFields = new Set();
    const addFieldLine = (fieldKey, lineText) => {
      if (!lineText) {
        return;
      }
      if (reportedFields.has(fieldKey)) {
        return;
      }
      reportedFields.add(fieldKey);
      lines.push(lineText);
    };
    const grossMatch = isSameAyalaAmount(result.totals.grossHourly, result.totals.grossEod);
    const epayMatch = isSameAyalaAmount(result.totals.epayHourly, result.totals.epayEod);
    const paymentDailyTotal = roundAyalaAmount(
      Object.values(result.paymentDailyBreakdown || {}).reduce((sum, value) => sum + value, 0)
    );
    const paymentTransactionTotal = roundAyalaAmount(
      Object.values(result.paymentTransactionBreakdown || {}).reduce((sum, value) => sum + value, 0)
    );
    const paymentMatch = getAyalaPaymentDiscrepancy(paymentDailyTotal, paymentTransactionTotal) <= AYALA_DISCREPANCY_TOLERANCE;
    const salesTotalMatch = isSameAyalaAmount(result.totals.salesTotalDaily, result.totals.salesTotalTransaction);
    const refundAmountMatch = isSameAyalaAmount(result.totals.refundAmountHourly, result.totals.refundAmountEod);
    const refundTxnMatch = result.totals.refundTxnHourly === result.totals.refundTxnEod;
    const noTrnMatch = result.eodSummary.noTrn === result.totals.hourlyTxnCount;
    const paymentOrder = [
      'CASH_SLS', 'OTHER_SLS', 'CHECK_SLS', 'GC_SLS', 'MASTERCARD_SLS', 'VISA_SLS',
      'AMEX_SLS', 'DINERS_SLS', 'JCB_SLS', 'GCASH_SLS', 'PAYMAYA_SLS', 'ALIPAY_SLS',
      'WECHAT_SLS', 'GRAB_SLS', 'FOODPANDA_SLS', 'MASTERDEBIT_SLS', 'VISADEBIT_SLS',
      'PAYPAL_SLS', 'ONLINE_SLS', 'OPEN_SALES', 'OPEN_SALES_2', 'OPEN_SALES_3',
      'OPEN_SALES_4', 'OPEN_SALES_5', 'OPEN_SALES_6', 'OPEN_SALES_7', 'OPEN_SALES_8',
      'OPEN_SALES_9', 'OPEN_SALES_10', 'OPEN_SALES_11', 'GC_EXCESS',
    ];
    const paymentDailyText = paymentOrder
      .map((field) => `${field} = ${formatAyalaAmount(result.paymentDailyBreakdown[field] || 0)}`)
      .join(', ');
    const paymentTransactionText = paymentOrder
      .map((field) => `${field === 'OTHER_SLS' ? 'OTHERSL_SLS' : field} = ${formatAyalaAmount(result.paymentTransactionBreakdown[field] || 0)}`)
      .join(', ');
    const paymentFieldDiffLines = paymentOrder
      .map((field) => {
        const dailyValue = roundAyalaAmount(result.paymentDailyBreakdown[field] || 0);
        const transactionValue = roundAyalaAmount(result.paymentTransactionBreakdown[field] || 0);
        const fieldDiscrepancy = getAyalaPaymentDiscrepancy(dailyValue, transactionValue);
        if (fieldDiscrepancy <= AYALA_DISCREPANCY_TOLERANCE) {
          return null;
        }
        const hourlyFieldName = field === 'OTHER_SLS' ? 'OTHERSL_SLS' : field;
        return {
          field,
          line: `Cross Validation, Discrepancy (${formatAyalaAmount(fieldDiscrepancy)}) ${field} DAILY (${formatAyalaAmount(dailyValue)}) and ${hourlyFieldName} TRANSACTION (${formatAyalaAmount(transactionValue)})`,
        };
      })
      .filter(Boolean);
    const hasMismatchGroups = result.mismatches.length > 0;
    const hasMissingGroups = result.missingInEod.length > 0 || result.missingInHourly.length > 0;
    const hasDuplicateTransactions = (result.duplicateTransactions || []).length > 0;
    const hasTotalDiscrepancy =
      !grossMatch
      || !epayMatch
      || !paymentMatch
      || !salesTotalMatch
      || !refundAmountMatch
      || !refundTxnMatch
      || !noTrnMatch
      || hasDuplicateTransactions
      || hasMismatchGroups
      || hasMissingGroups;
    const displayDates = result.eodSummary.trnDates && result.eodSummary.trnDates.length > 0
      ? result.eodSummary.trnDates.join(', ')
      : (result.eodSummary.trnDate || '-');
    const displayTerminals = result.eodSummary.terNos && result.eodSummary.terNos.length > 0
      ? result.eodSummary.terNos.join(', ')
      : (result.eodSummary.terNo || '-');
    const headerPrefix = `Transaction Date (${displayDates}), TERMINAL NO (${displayTerminals}), `;

    if (!hasTotalDiscrepancy) {
      lines.push(`${headerPrefix}No discrepancy found between CSV Hourly totals and EOD totals.`);
    } else {
      if (!grossMatch) {
        addFieldLine('GROSS_SLS',
          `${headerPrefix}Cross Validation, Discrepancy (${formatAyalaAmount(Math.abs(result.totals.grossEod - result.totals.grossHourly))}) `
          + `GROSS DAILY (${formatAyalaAmount(result.totals.grossEod)}) and GROSS TRANSACTION (${formatAyalaAmount(result.totals.grossHourly)})`
        );
      }
      if (!epayMatch) {
        addFieldLine('EPAY_SLS',
          `Cross Validation, Discrepancy (${formatAyalaAmount(Math.abs(result.totals.epayEod - result.totals.epayHourly))}) `
          + `EPAY DAILY (${formatAyalaAmount(result.totals.epayEod)}) and EPAY TRANSACTION (${formatAyalaAmount(result.totals.epayHourly)})`
        );
      }
      if (!paymentMatch) {
        addFieldLine('PAYMENT_TOTAL',
          `Cross Validation, Discrepancy (${formatAyalaAmount(Math.abs(paymentDailyTotal - paymentTransactionTotal))}) `
          + `PAYMENT DAILY (${formatAyalaAmount(paymentDailyTotal)}) : (${paymentDailyText}, ). and `
          + `PAYMENT TRANSACTION (${formatAyalaAmount(paymentTransactionTotal)}) : (${paymentTransactionText}, ).`
        );
      }
      paymentFieldDiffLines.forEach((item) => addFieldLine(item.field, item.line));
      if (!salesTotalMatch) {
        addFieldLine('SALES_TOTAL',
          `Cross Validation, Discrepancy (${formatAyalaAmount(Math.abs(result.totals.salesTotalDaily - result.totals.salesTotalTransaction))}) `
          + `SALES_TOTAL DAILY (${formatAyalaAmount(result.totals.salesTotalDaily)}) : `
          + `(NEW_GRNTOT = ${formatAyalaAmount(result.totals.newGrnTotEod)}, OLD_GRNTOT = ${formatAyalaAmount(result.totals.oldGrnTotEod)}, ). and `
          + `SALES_TOTAL TRANSACTION (${formatAyalaAmount(result.totals.salesTotalTransaction)}) : `
          + `(VAT_AMNT = ${formatAyalaAmount(result.totals.vatAmntHourly)}, VATABLE_SLS = ${formatAyalaAmount(result.totals.vatableSlsHourly)}, `
          + `NONVAT_SLS = ${formatAyalaAmount(result.totals.nonvatSlsHourly)}, VATEXEMPT_SLS = ${formatAyalaAmount(result.totals.vatexemptSlsHourly)}, ).`
        );
      }
      if (!noTrnMatch) {
        addFieldLine('NO_TRN',
          `Cross Validation, Discrepancy (${Math.abs(result.eodSummary.noTrn - result.totals.hourlyTxnCount)}) `
          + `NO_TRN (${result.eodSummary.noTrn}) not equal to total of transaction (${result.totals.hourlyTxnCount})`
        );
      }
      if (!refundAmountMatch) {
        addFieldLine('REFUND_AMT',
          `Cross Validation, Discrepancy (${formatAyalaAmount(Math.abs(result.totals.refundAmountEod - result.totals.refundAmountHourly))}) `
          + `REFUND DAILY (${formatAyalaAmount(result.totals.refundAmountEod)}) and REFUND TRANSACTION (${formatAyalaAmount(result.totals.refundAmountHourly)})`
        );
      }
      if (!refundTxnMatch) {
        addFieldLine('NO_REFUND',
          `Cross Validation, Discrepancy (${Math.abs(result.totals.refundTxnEod - result.totals.refundTxnHourly)}) `
          + `NO_REFUND (${result.totals.refundTxnEod}) not equal to refund transaction count (${result.totals.refundTxnHourly})`
        );
      }
      if (hasDuplicateTransactions) {
        lines.push(
          `Cross Validation, Discrepancy (${result.duplicateTransactions.length}) Duplicate TRANSACTION_NO detected in CSV Hourly.`
        );
        result.duplicateTransactions.forEach((dup) => {
          lines.push(
            `Duplicate TRANSACTION_NO (${dup.txnNo}), TERMINAL NO (${dup.terNo}), occurrences (${dup.count})`
          );
        });
      }

      if (result.missingInEod.length > 0) {
        lines.push(`Cross Validation, Discrepancy (${result.missingInEod.length}) Missing groups in EOD.`);
      }
      if (result.missingInHourly.length > 0) {
        lines.push(`Cross Validation, Discrepancy (${result.missingInHourly.length}) Missing groups in Hourly.`);
      }
      if (result.mismatches.length > 0) {
        result.mismatches.forEach((mismatch) => {
          mismatch.issues.forEach((issue) => {
            const numericMatch = issue.match(/^([A-Z0-9_]+) mismatch: hourly=([0-9.]+), eod=([0-9.]+)$/);
            if (numericMatch) {
              // Group-level numeric mismatch can be false-positive when moved refund/original pairs
              // are rendered in checker matrix mode. Use summary-level checks as source of truth.
              return;
            }
            lines.push(`Cross Validation, Discrepancy (1) ${issue}`);
          });
        });
      }
    }

    const terminalLines = (result.terminalSummaries || [])
      .sort((a, b) => String(a.terNo).localeCompare(String(b.terNo)))
      .map((terminal) => {
        const grossDiff = Math.abs(terminal.eodGross - terminal.hourlyGross);
        const epayDiff = Math.abs(terminal.eodEpay - terminal.hourlyEpay);
        const noTrnDiff = Math.abs(terminal.eodNoTrn - terminal.hourlyTxn);
        const grossText = isSameAyalaAmount(terminal.eodGross, terminal.hourlyGross)
          ? 'GROSS MATCH'
          : `GROSS DIFF ${formatAyalaAmount(grossDiff)}`;
        const epayText = isSameAyalaAmount(terminal.eodEpay, terminal.hourlyEpay)
          ? 'EPAY MATCH'
          : `EPAY DIFF ${formatAyalaAmount(epayDiff)}`;
        const noTrnText = terminal.eodNoTrn === terminal.hourlyTxn
          ? 'NO_TRN MATCH'
          : `NO_TRN DIFF ${noTrnDiff}`;
        const duplicateText = terminal.duplicateTxnCount > 0
          ? `, DUP_TRN ${terminal.duplicateTxnCount}`
          : '';
        return `TERMINAL ${terminal.terNo}: ${grossText}, ${epayText}, ${noTrnText}${duplicateText} (Hourly NO_TRN ${terminal.hourlyTxn}, EOD NO_TRN ${terminal.eodNoTrn})`;
      });
    if (terminalLines.length > 0) {
      lines.push('');
      lines.push('[Per Terminal Summary]');
      terminalLines.forEach((line) => lines.push(line));
    }

    return lines.join('\n');
  }

  function buildAyalaTransactionDiscrepancyReport(hourlyRecords) {
    const paymentTotalFields = [
      'CASH_SLS', 'OTHERSL_SLS', 'CHECK_SLS', 'GC_SLS', 'MASTERCARD_SLS', 'VISA_SLS',
      'AMEX_SLS', 'DINERS_SLS', 'JCB_SLS', 'GCASH_SLS', 'PAYMAYA_SLS', 'ALIPAY_SLS',
      'WECHAT_SLS', 'GRAB_SLS', 'FOODPANDA_SLS', 'MASTERDEBIT_SLS', 'VISADEBIT_SLS',
      'PAYPAL_SLS', 'ONLINE_SLS', 'OPEN_SALES', 'OPEN_SALES_2', 'OPEN_SALES_3',
      'OPEN_SALES_4', 'OPEN_SALES_5', 'OPEN_SALES_6', 'OPEN_SALES_7', 'OPEN_SALES_8',
      'OPEN_SALES_9', 'OPEN_SALES_10', 'OPEN_SALES_11', 'GC_EXCESS',
    ];
    const paymentSourceFields = ['VAT_AMNT', 'VATABLE_SLS', 'NONVAT_SLS', 'VATEXEMPT_SLS', 'SCHRGE_AMT', 'OTHER_SCHR', 'LOCAL_TAX'];
    const grossTotalFields = [
      'VAT_AMNT', 'VATABLE_SLS', 'NONVAT_SLS', 'VATEXEMPT_SLS', 'VATEXEMPT_AMNT',
      'LOCAL_TAX', 'PWD_DISC', 'SNRCIT_DISC', 'EMPLO_DISC', 'AYALA_DISC', 'STORE_DISC',
      'OTHER_DISC', 'SCHRGE_AMT', 'OTHER_SCHR',
    ];

    const formatTxnAmount = (value) => {
      const n = roundAyalaAmount(parseNumericAyala(value));
      if (Math.abs(n) < 0.005) {
        return '.00';
      }
      return n.toFixed(2);
    };

    const sorted = (hourlyRecords || [])
      .map((record) => ({
        txn: String(record.TRANSACTION_NO || '').trim(),
        ter: String(record.TER_NO || '').trim(),
        key: getAyalaTransactionCompositeKey(record),
        record,
      }))
      .filter((item) => item.txn && item.key && item.key !== '|')
      .sort((a, b) => {
        return compareAyalaTxnEntries(a.ter, a.txn, b.ter, b.txn);
      });

    const seen = new Set();
    const lines = [];
    sorted.forEach((item) => {
      if (seen.has(item.key)) {
        return;
      }
      seen.add(item.key);
      const record = item.record;
      const grossSls = parseNumericAyala(record.GROSS_SLS);
      const grossTotal = grossTotalFields.reduce((sum, field) => sum + parseNumericAyala(record[field]), 0);
      const grossDiscrepancy = getAyalaFlexibleDiscrepancy(grossSls, grossTotal, { allowOppositeSignSameAbs: true });
      const payment = paymentSourceFields.reduce((sum, field) => sum + getAyalaNumericFieldValue(record, field), 0);
      const paymentTotal = paymentTotalFields.reduce((sum, field) => sum + getAyalaAdjustedPaymentFieldValue(record, field), 0);
      const discrepancy = getAyalaPaymentDiscrepancy(payment, paymentTotal);
      const hasGrossIssue = grossDiscrepancy > AYALA_DISCREPANCY_TOLERANCE;
      const hasPaymentIssue = discrepancy > AYALA_DISCREPANCY_TOLERANCE;
      if (!hasGrossIssue && !hasPaymentIssue) return;

      const txNo = String(item.txn || '').padEnd(15, ' ');
      const terNo = String(record.TER_NO || '-');
      if (hasGrossIssue) {
        const grossTotalText = grossTotalFields
          .map((field) => `${field} = ${formatTxnAmount(record[field])}`)
          .join(', ');
        lines.push(
          `TRANSACTION NO (${txNo}), TERMINAL NO (${terNo}), `
          + `Discrepancy (${formatAyalaAmount(grossDiscrepancy)}) GROSS_SLS (${formatAyalaAmount(grossSls)}) and GROSS TOTAL (${formatAyalaAmount(grossTotal)}) :  `
          + `(${grossTotalText}, ).`
        );
      }

      if (hasPaymentIssue) {
        const paymentSourceText = paymentSourceFields
          .map((field) => `${field} = ${formatTxnAmount(getAyalaNumericFieldValue(record, field))}`)
          .join(', ');
        const paymentTotalText = paymentTotalFields
          .map((field) => `${field} = ${formatTxnAmount(getAyalaNumericFieldValue(record, field))}`)
          .join(', ');
        lines.push(
          `TRANSACTION NO (${txNo}), TERMINAL NO (${terNo}), `
          + `Discrepancy (${formatAyalaAmount(discrepancy)}) PAYMENT (${formatAyalaAmount(payment)}) : `
          + `(${paymentSourceText}, ). and PAYMENT TOTAL (${formatAyalaAmount(paymentTotal)}) :  `
          + `(${paymentTotalText}, ).`
        );
      }
    });

    if (lines.length === 0) {
      return 'No per-transaction discrepancy found.';
    }
    return lines.join('\n\n');
  }

  function checkRequiredAyalaHeaders(headers, requiredHeaders) {
    const headerSet = new Set(headers || []);
    return requiredHeaders.filter((header) => !headerSet.has(header));
  }

  async function readAyalaInputFile(file) {
    if (!file) {
      return { headers: [], records: [], parseError: 'No file selected.' };
    }
    const lowerName = file.name.toLowerCase();

    if (lowerName.endsWith('.xlsx')) {
      const buffer = await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => {
          const result = reader.result;
          if (result instanceof ArrayBuffer) {
            resolve(result);
            return;
          }
          reject(new Error('Failed to read Excel file.'));
        };
        reader.onerror = () => reject(new Error('Failed to read Excel file.'));
        reader.readAsArrayBuffer(file);
      });
      return parseAyalaTableFromWorkbook(buffer);
    }

    const text = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(typeof reader.result === 'string' ? reader.result : '');
      reader.onerror = () => reject(new Error('Failed to read text file.'));
      reader.readAsText(file);
    });
    return parseAyalaTableFromText(text);
  }

  async function readAyalaHourlyFiles(files) {
    const combinedHeaders = new Set();
    const combinedRecords = [];

    for (let i = 0; i < files.length; i += 1) {
      const file = files[i];
      const parsed = await readAyalaInputFile(file);
      if (parsed.parseError) {
        return { headers: [], records: [], parseError: `${file.name}: ${parsed.parseError}` };
      }
      (parsed.headers || []).forEach((header) => combinedHeaders.add(header));
      (parsed.records || []).forEach((record, recordIndex) => combinedRecords.push({
        ...record,
        __sourceFile: file.name,
        __sourceIndex: recordIndex,
      }));
    }

    return {
      headers: Array.from(combinedHeaders),
      records: combinedRecords,
    };
  }

  function getAyalaFilesSignature(files) {
    return (files || [])
      .map((file) => `${file.name}|${file.size}|${file.lastModified}`)
      .sort()
      .join('||');
  }

  function normalizeAyalaDateKey(raw) {
    const text = String(raw || '').trim();
    if (!text) return '';
    const isValidYmd = (y, m, d) => {
      const yy = Number.parseInt(String(y), 10);
      const mm = Number.parseInt(String(m), 10);
      const dd = Number.parseInt(String(d), 10);
      if (!Number.isFinite(yy) || !Number.isFinite(mm) || !Number.isFinite(dd)) return false;
      if (mm < 1 || mm > 12) return false;
      if (dd < 1 || dd > 31) return false;
      return true;
    };
    const yyyyMmDd = text.match(/^(\d{4})[-/](\d{2})[-/](\d{2})$/);
    if (yyyyMmDd) {
      if (!isValidYmd(yyyyMmDd[1], yyyyMmDd[2], yyyyMmDd[3])) return '';
      return `${yyyyMmDd[1]}-${yyyyMmDd[2]}-${yyyyMmDd[3]}`;
    }
    const compactYmd = text.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (compactYmd) {
      if (!isValidYmd(compactYmd[1], compactYmd[2], compactYmd[3])) return '';
      return `${compactYmd[1]}-${compactYmd[2]}-${compactYmd[3]}`;
    }
    const compactDmy = text.match(/^(\d{2})(\d{2})(\d{2})$/);
    if (compactDmy) {
      const y = `20${compactDmy[3]}`;
      const m = compactDmy[2];
      const d = compactDmy[1];
      if (!isValidYmd(y, m, d)) return '';
      return `${y}-${m}-${d}`;
    }
    return '';
  }

  function extractAyalaDateFromFileName(fileName) {
    const name = String(fileName || '').trim();
    if (!name) return '';
    const base = name.replace(/\.[^.]+$/, '');

    const ymd = base.match(/(20\d{2})[-_]?([01]\d)[-_]?([0-3]\d)/);
    if (ymd) {
      return normalizeAyalaDateKey(`${ymd[1]}-${ymd[2]}-${ymd[3]}`);
    }

    // Common hourly pattern: ...DDMMYYTTT_XXXXX (date + terminal + txn).
    const tailDmyWithTerminal = base.match(/(\d{2})(\d{2})(\d{2})\d{3}(?:_[0-9]+)?$/);
    if (tailDmyWithTerminal) {
      const parsed = normalizeAyalaDateKey(`${tailDmyWithTerminal[1]}${tailDmyWithTerminal[2]}${tailDmyWithTerminal[3]}`);
      if (parsed) return parsed;
    }

    const tailDmy = base.match(/(\d{2})(\d{2})(\d{2})(?:_[0-9]+)?$/);
    if (tailDmy) {
      const parsed = normalizeAyalaDateKey(`${tailDmy[1]}${tailDmy[2]}${tailDmy[3]}`);
      if (parsed) return parsed;
    }

    const anySix = Array.from(base.matchAll(/(\d{6})/g));
    if (anySix.length > 0) {
      // Prefer the first valid DDMMYY candidate from right-to-left.
      for (let i = anySix.length - 1; i >= 0; i -= 1) {
        const candidate = anySix[i][1];
        const parsed = normalizeAyalaDateKey(candidate);
        if (parsed) return parsed;
      }
    }
    return '';
  }

  function getAyalaRecordDateKey(record) {
    const trnDate = normalizeAyalaDateKey(record && record.TRN_DATE ? record.TRN_DATE : '');
    if (trnDate) return trnDate;
    const sourceFile = String(record && record.__sourceFile ? record.__sourceFile : '').trim();
    const fileDate = extractAyalaDateFromFileName(sourceFile);
    if (fileDate) return fileDate;
    return '';
  }

  function buildAyalaSourceFileDateMap(records) {
    const map = new Map();
    const byFile = new Map();
    (records || []).forEach((record) => {
      const sourceFile = String(record && record.__sourceFile ? record.__sourceFile : '').trim();
      if (!sourceFile) {
        return;
      }
      if (!byFile.has(sourceFile)) {
        byFile.set(sourceFile, []);
      }
      byFile.get(sourceFile).push(record);
    });

    byFile.forEach((fileRecords, sourceFile) => {
      let resolvedDate = '';
      for (let i = 0; i < fileRecords.length; i += 1) {
        const candidate = normalizeAyalaDateKey(fileRecords[i] && fileRecords[i].TRN_DATE ? fileRecords[i].TRN_DATE : '');
        if (candidate) {
          resolvedDate = candidate;
          break;
        }
      }
      if (!resolvedDate) {
        resolvedDate = extractAyalaDateFromFileName(sourceFile);
      }
      if (resolvedDate) {
        map.set(sourceFile, resolvedDate);
      }
    });

    return map;
  }

  function groupAyalaRecordsByDate(records, sourceFileDateMap) {
    const map = new Map();
    (records || []).forEach((record) => {
      const sourceFile = String(record && record.__sourceFile ? record.__sourceFile : '').trim();
      const mappedDate = sourceFile && sourceFileDateMap instanceof Map
        ? String(sourceFileDateMap.get(sourceFile) || '').trim()
        : '';
      const dateKey = mappedDate || getAyalaRecordDateKey(record) || 'UNKNOWN_DATE';
      if (!map.has(dateKey)) {
        map.set(dateKey, []);
      }
      map.get(dateKey).push(record);
    });
    return map;
  }

  function mergeAyalaEodRawPairGrids(grids) {
    const sourceGrids = (Array.isArray(grids) ? grids : []).filter((grid) => {
      return grid && Array.isArray(grid.rows) && grid.rows.length > 0;
    });
    if (sourceGrids.length === 0) {
      return { isMulti: false, terminals: [], rows: [] };
    }
    if (sourceGrids.length === 1) {
      return sourceGrids[0];
    }

    const summaryTotalFields = new Set(['GROSS_SLS', 'VAT_AMNT', 'VATABLE_SLS', 'REFUND_AMT']);
    const fieldOrder = [];
    const fieldSet = new Set();
    sourceGrids.forEach((grid) => {
      (grid.rows || []).forEach((row) => {
        const field = String(row && row.field ? row.field : '').trim();
        if (!field || fieldSet.has(field)) {
          return;
        }
        fieldSet.add(field);
        fieldOrder.push(field);
      });
    });

    const terminals = [];
    const terminalSeen = new Map();
    const valueMatrix = new Map();
    const appendValues = (field, values) => {
      if (!valueMatrix.has(field)) {
        valueMatrix.set(field, []);
      }
      valueMatrix.get(field).push(...values);
    };

    sourceGrids.forEach((grid, gridIndex) => {
      const rowsByField = new Map();
      (grid.rows || []).forEach((row) => {
        const field = String(row && row.field ? row.field : '').trim();
        if (field) {
          rowsByField.set(field, row);
        }
      });

      const valueCols = Math.max(
        1,
        ...(grid.rows || []).map((row) => {
          return Array.isArray(row && row.values) ? row.values.length : 0;
        })
      );
      const terRow = rowsByField.get('TER_NO');
      const terValues = Array.isArray(terRow && terRow.values) ? terRow.values : [];
      const labels = Array.from({ length: valueCols }, (_, idx) => {
        const baseLabel = String((grid.terminals || [])[idx] || terValues[idx] || `TER${idx + 1}`).trim() || `TER${idx + 1}`;
        const usedCount = terminalSeen.get(baseLabel) || 0;
        terminalSeen.set(baseLabel, usedCount + 1);
        if (usedCount === 0) {
          return baseLabel;
        }
        return `${baseLabel}_${gridIndex + 1}_${usedCount + 1}`;
      });
      terminals.push(...labels);

      fieldOrder.forEach((field) => {
        const row = rowsByField.get(field);
        const values = Array.from({ length: valueCols }, (_, idx) => {
          return String((row && Array.isArray(row.values) ? row.values[idx] : '') ?? '').trim();
        });
        appendValues(field, values);
      });
    });

    const rows = fieldOrder.map((field) => {
      const values = Array.isArray(valueMatrix.get(field)) ? valueMatrix.get(field) : [];
      let total = '';
      if (terminals.length > 1 && summaryTotalFields.has(field)) {
        total = formatAyalaAmount(values.reduce((sum, value) => sum + parseNumericAyala(value), 0));
      }
      return {
        field,
        values,
        total,
      };
    });

    return {
      isMulti: terminals.length > 1,
      terminals,
      rows,
    };
  }

  function buildAyalaEodRawPairsByDate(rawPairsByFile, sourceFileDateMap) {
    const byDate = new Map();
    if (!(rawPairsByFile instanceof Map)) {
      return byDate;
    }
    rawPairsByFile.forEach((grid, fileName) => {
      const sourceFile = String(fileName || '').trim();
      const mappedDate = sourceFile && sourceFileDateMap instanceof Map
        ? String(sourceFileDateMap.get(sourceFile) || '').trim()
        : '';
      const dateKey = mappedDate || extractAyalaDateFromFileName(sourceFile) || 'UNKNOWN_DATE';
      if (!byDate.has(dateKey)) {
        byDate.set(dateKey, []);
      }
      byDate.get(dateKey).push(grid);
    });
    return byDate;
  }

  async function readAyalaEodFiles(files) {
    const combinedHeaders = new Set();
    const combinedRecords = [];
    const rawPairsByFile = new Map();

    for (let i = 0; i < files.length; i += 1) {
      const file = files[i];
      const parsed = await readAyalaInputFile(file);
      if (parsed.parseError) {
        return { headers: [], records: [], rawPairsByFile: new Map(), parseError: `${file.name}: ${parsed.parseError}` };
      }
      (parsed.headers || []).forEach((header) => combinedHeaders.add(header));
      (parsed.records || []).forEach((record, recordIndex) => combinedRecords.push({
        ...record,
        __sourceFile: file.name,
        __sourceIndex: recordIndex,
      }));
      const rawPairs = await readAyalaRawPairs(file, parsed);
      rawPairsByFile.set(file.name, rawPairs);
    }

    return {
      headers: Array.from(combinedHeaders),
      records: combinedRecords,
      rawPairsByFile,
    };
  }

  function resetAyalaMergeState() {
    ayalaMergedHourlyData = null;
    ayalaMergedHourlySignature = '';
    ayalaHourlyAnalyzed = false;
  }

  function setAyalaHourlyFirstMode(isHourlyReady) {
    if (eodUploadCard) {
      eodUploadCard.hidden = !isHourlyReady;
    }
    if (eodCsvInput) {
      eodCsvInput.disabled = !isHourlyReady;
    }
    if (uploadAyalaBtn) {
      uploadAyalaBtn.hidden = !isHourlyReady;
      uploadAyalaBtn.disabled = !isHourlyReady;
    }
    if (mergeAyalaBtn) {
      mergeAyalaBtn.textContent = isHourlyReady ? 'Re-analyze Hourly' : 'Analyze Hourly';
    }
  }

  function formatMergeDisplayValue(field, rawValue) {
    const value = String(rawValue ?? '').trim();
    if (!value) {
      return '';
    }

    if (field === 'CCCODE') {
      const digitsOnly = value.replace(/[^0-9]/g, '');
      if (digitsOnly.length > 0) {
        const n = Number(digitsOnly);
        if (Number.isFinite(n)) {
          const exp = n.toExponential(14).replace('e', 'E');
          const match = exp.match(/^(.*)E([+-]?)(\d+)$/);
          if (match) {
            const sign = match[2] || '+';
            const paddedPow = match[3].padStart(3, '0');
            return `${match[1]}E${sign}${paddedPow}`;
          }
          return exp;
        }
      }
    }

    if (field === 'TRN_TIME') {
      const match = value.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
      if (match) {
        let hours = Number.parseInt(match[1], 10);
        const minutes = match[2];
        const seconds = match[3] || '00';
        const suffix = hours >= 12 ? 'PM' : 'AM';
        if (hours === 0) {
          hours = 12;
        } else if (hours > 12) {
          hours -= 12;
        }
        return `${hours}:${minutes}:${seconds} ${suffix}`;
      }
    }

    if (field === 'TER_NO') {
      const n = Number.parseInt(value, 10);
      if (Number.isFinite(n)) {
        return String(n);
      }
    }

    if (field === 'QTY' || field === 'QTY_SLD') {
      const n = Number.parseFloat(value);
      if (Number.isFinite(n)) {
        return Number.isInteger(n) ? String(n) : String(n);
      }
    }

    if (/^-?\d+(?:\.\d+)?$/.test(value)) {
      const n = Number.parseFloat(value);
      if (Number.isFinite(n)) {
        if (Number.isInteger(n)) {
          return String(n);
        }
        return String(n);
      }
    }

    return value;
  }

  function getAyalaRawMultiLineFieldValue(records, field) {
    const list = Array.isArray(records) ? records : [];
    if (list.length === 0) {
      return '';
    }
    const values = list
      .flatMap((record) => getAyalaRawFieldValues(record, field))
      .map((value) => String(value ?? '').trim())
      .filter((value) => value !== '');
    if (values.length === 0) {
      return '';
    }
    return values.join('\n');
  }

  function getAyalaRawInlineFieldValue(records, field) {
    const list = Array.isArray(records) ? records : [];
    if (list.length === 0) {
      return '';
    }
    const values = list
      .flatMap((record) => getAyalaRawFieldValues(record, field))
      .map((value) => String(value ?? '').trim())
      .filter((value) => value !== '');
    if (values.length === 0) {
      return '';
    }
    return values.join(' | ');
  }

  function buildAyalaHourlyMergeTable(hourlyRecords) {
    const orderedFields = [
      'CCCODE', 'MERCHANT_NAME', 'TRN_DATE', 'NO_TRN', 'CDATE', 'TRN_TIME', 'TER_NO', 'TRANSACTION_NO',
      'GROSS_SLS', 'VAT_AMNT', 'VATABLE_SLS', 'NONVAT_SLS', 'VATEXEMPT_SLS', 'VATEXEMPT_AMNT',
      'LOCAL_TAX', 'PWD_DISC', 'SNRCIT_DISC', 'EMPLO_DISC', 'AYALA_DISC', 'STORE_DISC', 'OTHER_DISC',
      'REFUND_AMT', 'SCHRGE_AMT', 'OTHER_SCHR', 'CASH_SLS', 'CARD_SLS', 'EPAY_SLS', 'DCARD_SLS',
      'OTHERSL_SLS', 'CHECK_SLS', 'GC_SLS', 'MASTERCARD_SLS', 'VISA_SLS', 'AMEX_SLS', 'DINERS_SLS',
      'JCB_SLS', 'GCASH_SLS', 'PAYMAYA_SLS', 'ALIPAY_SLS', 'WECHAT_SLS', 'GRAB_SLS', 'FOODPANDA_SLS',
      'MASTERDEBIT_SLS', 'VISADEBIT_SLS', 'PAYPAL_SLS', 'ONLINE_SLS', 'OPEN_SALES', 'OPEN_SALES_2',
      'OPEN_SALES_3', 'OPEN_SALES_4', 'OPEN_SALES_5', 'OPEN_SALES_6', 'OPEN_SALES_7', 'OPEN_SALES_8',
      'OPEN_SALES_9', 'OPEN_SALES_10', 'OPEN_SALES_11', 'GC_EXCESS', 'MOBILE_NO', 'NO_CUST',
      'TRN_TYPE', 'SLS_FLAG', 'VAT_PCT', 'QTY_SLD', 'QTY', 'ITEMCODE', 'PRICE', 'LDISC',
    ];

    const normalizedRows = hourlyRecords
      .map((record) => ({
        txn: String(record.TRANSACTION_NO || '').trim(),
        ter: String(record.TER_NO || '').trim(),
        key: getAyalaTransactionCompositeKey(record),
        record,
      }))
      .sort((a, b) => {
        return compareAyalaTxnEntries(a.ter, a.txn, b.ter, b.txn);
      });

    const firstOnlyFields = new Set(['CCCODE', 'MERCHANT_NAME', 'TRN_DATE']);
    const sumStartIndex = orderedFields.indexOf('GROSS_SLS');
    const sumEndIndex = orderedFields.indexOf('LDISC');

    const txnMap = new Map();
    const txnRowsMap = new Map();
    normalizedRows.forEach((row) => {
      if (!row.txn || !row.key || row.key === '|') {
        return;
      }
      if (!txnRowsMap.has(row.key)) {
        txnRowsMap.set(row.key, []);
      }
      txnRowsMap.get(row.key).push(row.record);
      if (txnMap.has(row.key)) {
        return;
      }
      txnMap.set(row.key, row.record);
    });

    const firstTxnKeyByFile = new Map();
    const txnCountByFile = new Map();
    txnMap.forEach((record) => {
      const sourceFile = String(record.__sourceFile || '').trim();
      const terNo = String(record.TER_NO || '').trim();
      const txnNumber = Number.parseInt(String(record.TRANSACTION_NO || '').trim(), 10);
      const txnKey = getAyalaTransactionCompositeKey(record);
      if (!sourceFile || !Number.isFinite(txnNumber)) {
        return;
      }
      if (!firstTxnKeyByFile.has(sourceFile)) {
        firstTxnKeyByFile.set(sourceFile, txnKey);
      } else {
        const currentFirstKey = String(firstTxnKeyByFile.get(sourceFile) || '');
        const [firstTer, firstTxn] = currentFirstKey.split('|');
        if (compareAyalaTxnEntries(terNo, String(txnNumber), String(firstTer || ''), String(firstTxn || '')) < 0) {
          firstTxnKeyByFile.set(sourceFile, txnKey);
        }
      }
      if (!txnCountByFile.has(sourceFile)) {
        txnCountByFile.set(sourceFile, new Set());
      }
      txnCountByFile.get(sourceFile).add(txnKey);
    });

    const sortedExistingTxns = Array.from(txnMap.keys()).sort((a, b) => {
      const [aTer, aTxn] = String(a).split('|');
      const [bTer, bTxn] = String(b).split('|');
      return compareAyalaTxnEntries(aTer || '', aTxn || '', bTer || '', bTxn || '');
    });
    let horizontalTxns = sortedExistingTxns;

    // Keep fixed merge row layout up to LDISC (manual checker format).

    const lines = [];

    orderedFields.forEach((field, fieldIndex) => {
      const isProductRawField = field === 'QTY' || field === 'ITEMCODE' || field === 'PRICE' || field === 'LDISC';
      const shouldSum = sumStartIndex >= 0
        && sumEndIndex >= 0
        && fieldIndex >= sumStartIndex
        && fieldIndex <= sumEndIndex
        && !isProductRawField;
      const values = horizontalTxns.map((txn) => {
        const record = txnMap.get(txn);
        if (!record) {
          return '';
        }
        const sourceFile = String(record.__sourceFile || '').trim();
        const currentTxnKey = getAyalaTransactionCompositeKey(record);
        const isFirstTxnForFile = sourceFile && firstTxnKeyByFile.get(sourceFile) === currentTxnKey;
        if (firstOnlyFields.has(field)) {
          if (!isFirstTxnForFile) {
            return '';
          }
        }
        if (field === 'NO_TRN') {
          if (isFirstTxnForFile) {
            const txnCount = txnCountByFile.has(sourceFile) ? txnCountByFile.get(sourceFile).size : 0;
            return String(txnCount || '');
          }
          return '';
        }
        if (isProductRawField) {
          const txnRows = txnRowsMap.get(txn) || [];
          return getAyalaRawMultiLineFieldValue(txnRows, field);
        }
        return formatMergeDisplayValue(field, getAyalaRawFieldValue(record, field));
      });
      const numericValues = values.map((value) => {
        if (!isProductRawField) {
          return parseNumericAyala(value);
        }
        return String(value || '')
          .split(/\r?\n/)
          .reduce((acc, item) => acc + parseNumericAyala(item), 0);
      });
      const sum = shouldSum
        ? formatAyalaAmount(numericValues.reduce((acc, n) => acc + n, 0))
        : '';

      lines.push([field, ...values, sum].join('\t'));
    });

    return lines.join('\n');
  }

  function buildAyalaHourlyAnalysisReport(hourlyRecords) {
    const normalizedRows = hourlyRecords
      .map((record) => ({
        txn: String(record.TRANSACTION_NO || '').trim(),
        ter: String(record.TER_NO || '').trim(),
        key: getAyalaTransactionCompositeKey(record),
        record,
      }))
      .filter((row) => row.txn && row.key && row.key !== '|')
      .sort((a, b) => {
        return compareAyalaTxnEntries(a.ter, a.txn, b.ter, b.txn);
      });

    const txnMap = new Map();
    normalizedRows.forEach((row) => {
      if (!txnMap.has(row.key)) {
        txnMap.set(row.key, row.record);
      }
    });

    const uniqueRecords = Array.from(txnMap.entries())
      .sort((a, b) => {
        const [aTer, aTxn] = String(a[0]).split('|');
        const [bTer, bTxn] = String(b[0]).split('|');
        return compareAyalaTxnEntries(aTer || '', aTxn || '', bTer || '', bTxn || '');
      })
      .map((entry) => entry[1]);

    const paymentFields = [
      'CASH_SLS', 'OTHERSL_SLS', 'CHECK_SLS', 'GC_SLS', 'MASTERCARD_SLS', 'VISA_SLS',
      'AMEX_SLS', 'DINERS_SLS', 'JCB_SLS', 'GCASH_SLS', 'PAYMAYA_SLS', 'ALIPAY_SLS',
      'WECHAT_SLS', 'GRAB_SLS', 'FOODPANDA_SLS', 'MASTERDEBIT_SLS', 'VISADEBIT_SLS',
      'PAYPAL_SLS', 'ONLINE_SLS', 'OPEN_SALES', 'OPEN_SALES_2', 'OPEN_SALES_3',
      'OPEN_SALES_4', 'OPEN_SALES_5', 'OPEN_SALES_6', 'OPEN_SALES_7', 'OPEN_SALES_8',
      'OPEN_SALES_9', 'OPEN_SALES_10', 'OPEN_SALES_11', 'GC_EXCESS',
    ];

    const paymentBreakdown = {};
    paymentFields.forEach((field) => {
      paymentBreakdown[field] = 0;
    });

    let gross = 0;
    let epay = 0;
    let vatAmnt = 0;
    let vatableSls = 0;
    let nonvatSls = 0;
    let vatexemptSls = 0;
    let refundAmt = 0;
    let refundCount = 0;

    uniqueRecords.forEach((record) => {
      gross += parseNumericAyala(record.GROSS_SLS);
      epay += parseNumericAyala(record.EPAY_SLS);
      vatAmnt += parseNumericAyala(record.VAT_AMNT);
      vatableSls += parseNumericAyala(record.VATABLE_SLS);
      nonvatSls += parseNumericAyala(record.NONVAT_SLS);
      vatexemptSls += parseNumericAyala(record.VATEXEMPT_SLS);

      const recordRefund = parseNumericAyala(record.REFUND_AMT);
      refundAmt += recordRefund;
      const slsFlag = String(record.SLS_FLAG || '').trim().toUpperCase();
      if (recordRefund > 0 || slsFlag === 'R' || slsFlag === 'REFUND') {
        refundCount += 1;
      }

      paymentFields.forEach((field) => {
        paymentBreakdown[field] += getAyalaAdjustedPaymentFieldValue(record, field);
      });
    });

    const salesTotal = vatAmnt + vatableSls + nonvatSls + vatexemptSls;
    const paymentTotal = paymentFields.reduce((sum, field) => sum + paymentBreakdown[field], 0);
    const strans = uniqueRecords.length > 0 ? String(uniqueRecords[0].TRANSACTION_NO || '') : '-';
    const etrans = uniqueRecords.length > 0 ? String(uniqueRecords[uniqueRecords.length - 1].TRANSACTION_NO || '') : '-';
    const trnDates = Array.from(new Set(
      uniqueRecords.map((record) => String(record.TRN_DATE || '').trim()).filter(Boolean)
    ));
    const terNos = Array.from(new Set(
      uniqueRecords.map((record) => String(record.TER_NO || '').trim()).filter(Boolean)
    ));
    const trnDate = trnDates.length > 0 ? trnDates.join(', ') : '-';
    const terNo = terNos.length > 0 ? terNos.join(', ') : '-';
    const perTerminalMap = new Map();
    uniqueRecords.forEach((record) => {
      const ter = String(record.TER_NO || '').trim() || '-';
      if (!perTerminalMap.has(ter)) {
        perTerminalMap.set(ter, {
          txnCount: 0,
          minTxn: null,
          maxTxn: null,
          gross: 0,
          epay: 0,
        });
      }
      const terminal = perTerminalMap.get(ter);
      terminal.txnCount += 1;
      terminal.gross += parseNumericAyala(record.GROSS_SLS);
      terminal.epay += parseNumericAyala(record.EPAY_SLS);
      const txnNum = Number.parseInt(String(record.TRANSACTION_NO || '').trim(), 10);
      if (Number.isFinite(txnNum)) {
        if (terminal.minTxn === null || txnNum < terminal.minTxn) terminal.minTxn = txnNum;
        if (terminal.maxTxn === null || txnNum > terminal.maxTxn) terminal.maxTxn = txnNum;
      }
    });
    const paymentText = paymentFields
      .map((field) => `${field} = ${formatAyalaAmount(paymentBreakdown[field] || 0)}`)
      .join(', ');

    const lines = [];
    lines.push(
      `Transaction Date (${trnDate}), TERMINAL NO (${terNo}), `
      + 'CSV Hourly files analyzed successfully.'
    );
    lines.push(`NO_TRN (${uniqueRecords.length})`);
    lines.push(`STRANS (${strans})`);
    lines.push(`ETRANS (${etrans})`);
    lines.push(`GROSS TRANSACTION (${formatAyalaAmount(gross)})`);
    lines.push(`EPAY TRANSACTION (${formatAyalaAmount(epay)})`);
    lines.push(
      `PAYMENT TRANSACTION (${formatAyalaAmount(paymentTotal)}) : `
      + `(${paymentText}, ).`
    );
    lines.push(
      `SALES_TOTAL TRANSACTION (${formatAyalaAmount(salesTotal)}) : `
      + `(VAT_AMNT = ${formatAyalaAmount(vatAmnt)}, VATABLE_SLS = ${formatAyalaAmount(vatableSls)}, `
      + `NONVAT_SLS = ${formatAyalaAmount(nonvatSls)}, VATEXEMPT_SLS = ${formatAyalaAmount(vatexemptSls)}, ).`
    );
    lines.push(`REFUND TRANSACTION (${formatAyalaAmount(refundAmt)})`);
    lines.push(`NO_REFUND (${refundCount})`);
    if (perTerminalMap.size > 1) {
      lines.push('');
      lines.push('[Per Terminal Hourly]');
      Array.from(perTerminalMap.entries())
        .sort((a, b) => String(a[0]).localeCompare(String(b[0])))
        .forEach(([ter, value]) => {
          const minTxn = value.minTxn === null ? '-' : String(value.minTxn);
          const maxTxn = value.maxTxn === null ? '-' : String(value.maxTxn);
          lines.push(
            `TERMINAL ${ter}: NO_TRN (${value.txnCount}), STRANS (${minTxn}), ETRANS (${maxTxn}), `
            + `GROSS (${formatAyalaAmount(value.gross)}), EPAY (${formatAyalaAmount(value.epay)})`
          );
        });
    }
    lines.push('');
    lines.push('Next step: choose EOD file and click Upload for cross validation.');
    return lines.join('\n');
  }

  function buildAyalaCheckerWorkbook(hourlyRecords, eodRecords, eodRawPairs) {
    const orderedFields = [
      'CCCODE', 'MERCHANT_NAME', 'TRN_DATE', 'NO_TRN', 'CDATE', 'TRN_TIME', 'TER_NO', 'TRANSACTION_NO',
      'GROSS_SLS', 'VAT_AMNT', 'VATABLE_SLS', 'NONVAT_SLS', 'VATEXEMPT_SLS', 'VATEXEMPT_AMNT',
      'LOCAL_TAX', 'PWD_DISC', 'SNRCIT_DISC', 'EMPLO_DISC', 'AYALA_DISC', 'STORE_DISC', 'OTHER_DISC',
      'REFUND_AMT', 'SCHRGE_AMT', 'OTHER_SCHR', 'CASH_SLS', 'CARD_SLS', 'EPAY_SLS', 'DCARD_SLS',
      'OTHERSL_SLS', 'CHECK_SLS', 'GC_SLS', 'MASTERCARD_SLS', 'VISA_SLS', 'AMEX_SLS', 'DINERS_SLS',
      'JCB_SLS', 'GCASH_SLS', 'PAYMAYA_SLS', 'ALIPAY_SLS', 'WECHAT_SLS', 'GRAB_SLS', 'FOODPANDA_SLS',
      'MASTERDEBIT_SLS', 'VISADEBIT_SLS', 'PAYPAL_SLS', 'ONLINE_SLS', 'OPEN_SALES', 'OPEN_SALES_2',
      'OPEN_SALES_3', 'OPEN_SALES_4', 'OPEN_SALES_5', 'OPEN_SALES_6', 'OPEN_SALES_7', 'OPEN_SALES_8',
      'OPEN_SALES_9', 'OPEN_SALES_10', 'OPEN_SALES_11', 'GC_EXCESS', 'MOBILE_NO', 'NO_CUST',
      'TRN_TYPE', 'SLS_FLAG', 'VAT_PCT', 'QTY_SLD', 'QTY', 'ITEMCODE', 'PRICE', 'LDISC',
    ];
    const firstOnlyFields = new Set(['CCCODE', 'MERCHANT_NAME', 'TRN_DATE']);
    const productRawFields = new Set(['QTY', 'ITEMCODE', 'PRICE', 'LDISC']);
    const productFieldOrder = ['QTY', 'ITEMCODE', 'PRICE', 'LDISC'];
    const sumStartIndex = orderedFields.indexOf('GROSS_SLS');
    const sumEndIndex = orderedFields.indexOf('LDISC');

    const normalizedRows = hourlyRecords
      .map((record) => ({
        txn: String(record.TRANSACTION_NO || '').trim(),
        ter: String(record.TER_NO || '').trim(),
        key: getAyalaTransactionCompositeKey(record),
        record,
      }))
      .filter((row) => row.txn && row.key && row.key !== '|')
      .sort((a, b) => {
        return compareAyalaTxnEntries(a.ter, a.txn, b.ter, b.txn);
      });

    const transactionEntries = normalizedRows
      .filter((row) => row.txn && row.key && row.key !== '|')
      .map((row) => ({
        key: row.key,
        txn: row.txn,
        ter: row.ter || '-',
        txnNum: Number.parseInt(row.txn, 10),
        record: row.record,
      }));

    const txnMap = new Map();
    const txnRowsMap = new Map();
    normalizedRows.forEach((row) => {
      if (!txnMap.has(row.key)) {
        txnMap.set(row.key, row.record);
      }
      if (!row.key || row.key === '|') {
        return;
      }
      if (!txnRowsMap.has(row.key)) {
        txnRowsMap.set(row.key, []);
      }
      txnRowsMap.get(row.key).push(row.record);
    });
    const getTxnProductLineValue = (txnKey, field, lineIndex = 0) => {
      const txnRows = txnRowsMap.get(txnKey) || [];
      const values = txnRows
        .flatMap((record) => getAyalaRawFieldValues(record, field))
        .map((value) => String(value ?? '').trim())
        .filter((value) => value !== '');
      const raw = values[lineIndex] || '';
      if (!raw) {
        return '';
      }
      return formatMergeDisplayValue(field, raw);
    };
    const maxProductLineCount = Math.max(
      1,
      ...Array.from(txnRowsMap.keys()).map((txnKey) => {
        return Math.max(
          ...productFieldOrder.map((field) => {
            const txnRows = txnRowsMap.get(txnKey) || [];
            return txnRows
              .flatMap((record) => getAyalaRawFieldValues(record, field))
              .map((value) => String(value ?? '').trim())
              .filter((value) => value !== '')
              .length;
          }),
          1
        );
      })
    );

    const firstTxnKeyByFile = new Map();
    const txnCountByFile = new Map();
    txnMap.forEach((record) => {
      const sourceFile = String(record.__sourceFile || '').trim();
      const terNo = String(record.TER_NO || '').trim();
      const txnNumber = Number.parseInt(String(record.TRANSACTION_NO || '').trim(), 10);
      const txnKey = getAyalaTransactionCompositeKey(record);
      if (!sourceFile || !Number.isFinite(txnNumber)) {
        return;
      }
      if (!firstTxnKeyByFile.has(sourceFile)) {
        firstTxnKeyByFile.set(sourceFile, txnKey);
      } else {
        const currentFirstKey = String(firstTxnKeyByFile.get(sourceFile) || '');
        const [firstTer, firstTxn] = currentFirstKey.split('|');
        if (compareAyalaTxnEntries(terNo, String(txnNumber), String(firstTer || ''), String(firstTxn || '')) < 0) {
          firstTxnKeyByFile.set(sourceFile, txnKey);
        }
      }
      if (!txnCountByFile.has(sourceFile)) {
        txnCountByFile.set(sourceFile, new Set());
      }
      txnCountByFile.get(sourceFile).add(txnKey);
    });

    const sortedExistingTxns = Array.from(txnMap.keys()).sort((a, b) => {
      const [aTer, aTxn] = String(a).split('|');
      const [bTer, bTxn] = String(b).split('|');
      return compareAyalaTxnEntries(aTer || '', aTxn || '', bTer || '', bTxn || '');
    });
    let horizontalTxns = sortedExistingTxns;

    const eodCombinedValues = buildAyalaCombinedFieldValues(eodRecords);
    const eodGrid = (eodRawPairs && typeof eodRawPairs === 'object' && Array.isArray(eodRawPairs.rows))
      ? eodRawPairs
      : { isMulti: false, terminals: [], rows: [] };
    const eodAlias = {
      OTHERSL_SLS: 'OTHER_SLS',
    };

    const aoa = [];
    const formulaRows = [];
    const comparisonRows = [];
    const firstTxnColIndex = 1; // column B in sheet (0-based index)
    const lastTxnColIndex = horizontalTxns.length; // B..(B+n-1), 0-based
    const hourlyTerminalOrder = Array.from(new Set(horizontalTxns.map((key) => String(key).split('|')[0] || '-')))
      .sort((a, b) => compareAyalaTxnEntries(a, '0', b, '0'));
    const hasMultiHourlyTerminal = hourlyTerminalOrder.length > 1;
    const terminalRangeMap = new Map();
    horizontalTxns.forEach((key, idx) => {
      const terNo = String(key).split('|')[0] || '-';
      const colIndex = firstTxnColIndex + idx;
      if (!terminalRangeMap.has(terNo)) {
        terminalRangeMap.set(terNo, { start: colIndex, end: colIndex });
      } else {
        terminalRangeMap.get(terNo).end = colIndex;
      }
    });

    const fieldCopyColIndex = horizontalTxns.length + 1;
    const hourlyTerminalStartColIndex = fieldCopyColIndex + 1;
    const hourlyTerminalCount = hasMultiHourlyTerminal ? hourlyTerminalOrder.length : 0;
    const sumColIndex = hourlyTerminalStartColIndex + hourlyTerminalCount;
    const separatorColIndex = sumColIndex + 1;
    const eodFieldColIndex = separatorColIndex + 1;
    const eodFirstValueColIndex = eodFieldColIndex + 1;
    const eodTerminalCount = eodGrid.isMulti ? eodGrid.terminals.length : 1;
    const eodTotalColIndex = eodGrid.isMulti ? (eodFirstValueColIndex + eodTerminalCount) : -1;
    const eodBlockEndColIndex = eodGrid.isMulti && eodTotalColIndex >= 0 ? eodTotalColIndex : eodFirstValueColIndex;
    const compareHighlightFields = new Set(['GROSS_SLS', 'VAT_AMNT', 'VATABLE_SLS', 'REFUND_AMT']);

    // Detect refund transactions and match to original transaction (same TER_NO and same absolute gross).
    const getRefundDerivedAmounts = (entry) => {
      const gross = parseNumericAyala(entry.record.GROSS_SLS);
      const refundAmt = Math.abs(parseNumericAyala(entry.record.REFUND_AMT));
      const vat = Math.abs(parseNumericAyala(entry.record.VAT_AMNT));
      const vatable = Math.abs(parseNumericAyala(entry.record.VATABLE_SLS));
      const nonvat = Math.abs(parseNumericAyala(entry.record.NONVAT_SLS));
      const vatexempt = Math.abs(parseNumericAyala(entry.record.VATEXEMPT_SLS));
      const schrge = Math.abs(parseNumericAyala(entry.record.SCHRGE_AMT));
      const otherSchr = Math.abs(parseNumericAyala(entry.record.OTHER_SCHR));
      const localTax = Math.abs(parseNumericAyala(entry.record.LOCAL_TAX));
      // Fallback for refund formats where gross is zero but components are populated.
      const componentBasedGross = roundAyalaAmount(vat + vatable + nonvat + vatexempt + schrge + otherSchr + localTax);
      const grossAbs = Math.abs(gross);
      const derivedGross = grossAbs > 0
        ? grossAbs
        : (refundAmt > 0 ? refundAmt : componentBasedGross);
      return {
        gross: derivedGross,
        vat,
        vatable,
      };
    };
    const isRefundRecord = (entry) => {
      const gross = parseNumericAyala(entry.record.GROSS_SLS);
      const refundAmt = parseNumericAyala(entry.record.REFUND_AMT);
      const slsFlag = String(entry.record.SLS_FLAG || '').trim().toUpperCase();
      const trnType = String(entry.record.TRN_TYPE || '').trim().toUpperCase();
      const derived = getRefundDerivedAmounts(entry);
      return (
        gross < 0
        || refundAmt > 0
        || slsFlag === 'R'
        || slsFlag === 'REFUND'
        || trnType === 'R'
        || trnType === 'REFUND'
        || derived.gross > 0 && refundAmt > 0
      );
    };
    const refundEntries = transactionEntries.filter((entry) => isRefundRecord(entry));
    const originalEntries = transactionEntries.filter((entry) => !isRefundRecord(entry));
    const usedOriginalKeys = new Set();
    const refundMatches = [];
    const unmatchedRefunds = [];
    refundEntries.forEach((refund) => {
      const refundDerived = getRefundDerivedAmounts(refund);
      const refundItem = String(refund.record.ITEMCODE || '').trim().toUpperCase();
      const refundPrice = Math.abs(parseNumericAyala(refund.record.PRICE));
      const refundQty = Math.abs(parseNumericAyala(refund.record.QTY));
      const refundGross = Math.abs(parseNumericAyala(refund.record.GROSS_SLS));
      const refundVatable = Math.abs(parseNumericAyala(refund.record.VATABLE_SLS));
      const refundVat = Math.abs(parseNumericAyala(refund.record.VAT_AMNT));

      const scoredCandidates = originalEntries
        .filter((orig) => {
          if (usedOriginalKeys.has(orig.key)) {
            return false;
          }
          if ((orig.ter || '-') !== (refund.ter || '-')) {
            return false;
          }
          if (Number.isFinite(orig.txnNum) && Number.isFinite(refund.txnNum)) {
            return orig.txnNum <= refund.txnNum;
          }
          return true;
        })
        .map((orig) => {
          const origItem = String(orig.record.ITEMCODE || '').trim().toUpperCase();
          const origPrice = Math.abs(parseNumericAyala(orig.record.PRICE));
          const origQty = Math.abs(parseNumericAyala(orig.record.QTY));
          const origGross = Math.abs(parseNumericAyala(orig.record.GROSS_SLS));
          const origVatable = Math.abs(parseNumericAyala(orig.record.VATABLE_SLS));
          const origVat = Math.abs(parseNumericAyala(orig.record.VAT_AMNT));
          let score = 0;

          if (refundItem && origItem && refundItem === origItem) {
            score += 100;
          }
          if (refundPrice > 0 && isSameAyalaAmount(refundPrice, origPrice)) {
            score += 60;
          }
          if (refundQty > 0 && isSameAyalaAmount(refundQty, origQty)) {
            score += 20;
          }
          if (refundGross > 0 && isSameAyalaAmount(refundGross, origGross)) {
            score += 40;
          } else if (refundDerived.gross > 0 && isSameAyalaAmount(refundDerived.gross, origGross)) {
            score += 30;
          }
          if (refundVatable > 0 && isSameAyalaAmount(refundVatable, origVatable)) {
            score += 35;
          }
          if (refundVat > 0 && isSameAyalaAmount(refundVat, origVat)) {
            score += 25;
          }

          const txnGap = Number.isFinite(orig.txnNum) && Number.isFinite(refund.txnNum)
            ? Math.abs(refund.txnNum - orig.txnNum)
            : 999999;
          return { orig, score, txnGap };
        })
        .sort((a, b) => {
          if (b.score !== a.score) {
            return b.score - a.score;
          }
          return a.txnGap - b.txnGap;
        });

      const best = scoredCandidates[0];
      if (!best || best.score < 20) {
        unmatchedRefunds.push(refund);
        return;
      }

      usedOriginalKeys.add(best.orig.key);
      refundMatches.push({ refund, original: best.orig });
    });

    const movedTxnKeys = new Set();
    // Move only paired transactions (original + refund) out of main hourly matrix.
    if (AYALA_MOVE_PAIRS_ENABLED) {
      refundMatches.forEach((pair) => {
        movedTxnKeys.add(pair.refund.key);
        movedTxnKeys.add(pair.original.key);
      });
    }
    const hasRefundScenario = AYALA_MOVE_PAIRS_ENABLED && refundEntries.length > 0;
    const adjustmentStartIndex = orderedFields.indexOf('GROSS_SLS');
    const adjustmentEndIndex = orderedFields.indexOf('OTHER_SCHR');
    const adjustmentFields = (adjustmentStartIndex >= 0 && adjustmentEndIndex >= adjustmentStartIndex)
      ? orderedFields.slice(adjustmentStartIndex, adjustmentEndIndex + 1)
      : ['GROSS_SLS', 'VAT_AMNT', 'VATABLE_SLS', 'NONVAT_SLS', 'VATEXEMPT_SLS', 'VATEXEMPT_AMNT', 'LOCAL_TAX', 'PWD_DISC', 'SNRCIT_DISC', 'EMPLO_DISC', 'AYALA_DISC', 'STORE_DISC', 'OTHER_DISC', 'REFUND_AMT', 'SCHRGE_AMT', 'OTHER_SCHR'];
    const movedOriginalGrossTotal = AYALA_MOVE_PAIRS_ENABLED
      ? roundAyalaAmount(
        refundMatches.reduce((acc, pair) => acc + parseNumericAyala(pair.original ? pair.original.record.GROSS_SLS : 0), 0)
      )
      : 0;
    const movedRefundTotal = AYALA_MOVE_PAIRS_ENABLED
      ? roundAyalaAmount(
        refundMatches.reduce((acc, pair) => acc + Math.abs(parseNumericAyala(pair.refund.record.REFUND_AMT)), 0)
      )
      : 0;
    const paymentAdjustmentFields = [
      'CASH_SLS', 'OTHERSL_SLS', 'CHECK_SLS', 'GC_SLS', 'MASTERCARD_SLS', 'VISA_SLS',
      'AMEX_SLS', 'DINERS_SLS', 'JCB_SLS', 'GCASH_SLS', 'PAYMAYA_SLS', 'ALIPAY_SLS',
      'WECHAT_SLS', 'GRAB_SLS', 'FOODPANDA_SLS', 'MASTERDEBIT_SLS', 'VISADEBIT_SLS',
      'PAYPAL_SLS', 'ONLINE_SLS', 'OPEN_SALES', 'OPEN_SALES_2', 'OPEN_SALES_3',
      'OPEN_SALES_4', 'OPEN_SALES_5', 'OPEN_SALES_6', 'OPEN_SALES_7', 'OPEN_SALES_8',
      'OPEN_SALES_9', 'OPEN_SALES_10', 'OPEN_SALES_11', 'GC_EXCESS',
    ];
    const movedPaymentAdjustByField = {};
    paymentAdjustmentFields.forEach((field) => {
      movedPaymentAdjustByField[field] = 0;
    });
    if (AYALA_MOVE_PAIRS_ENABLED && refundMatches.length > 0) {
      refundMatches.forEach((pair) => {
        paymentAdjustmentFields.forEach((field) => {
          movedPaymentAdjustByField[field] += getAyalaAdjustedPaymentFieldValue(pair.refund.record, field)
            + getAyalaAdjustedPaymentFieldValue(pair.original ? pair.original.record : null, field);
        });
      });
    }

    const fieldRowMap = new Map();
    const productLineRowMap = new Map();
    orderedFields.forEach((field, fieldIndex) => {
      const shouldSum = sumStartIndex >= 0
        && sumEndIndex >= 0
        && fieldIndex >= sumStartIndex
        && fieldIndex <= sumEndIndex
        && !productRawFields.has(field);
      const values = horizontalTxns.map((txn) => {
        if (movedTxnKeys.has(txn)) {
          return '';
        }
        const record = txnMap.get(txn);
        if (!record) {
          return '';
        }
        const sourceFile = String(record.__sourceFile || '').trim();
        const currentTxnKey = getAyalaTransactionCompositeKey(record);
        const isFirstTxnForFile = sourceFile && firstTxnKeyByFile.get(sourceFile) === currentTxnKey;
        if (firstOnlyFields.has(field)) {
          if (!isFirstTxnForFile) {
            return '';
          }
        }
        if (field === 'NO_TRN') {
          if (isFirstTxnForFile) {
            const txnCount = txnCountByFile.has(sourceFile) ? txnCountByFile.get(sourceFile).size : 0;
            return String(txnCount || '');
          }
          return '';
        }
        if (productRawFields.has(field)) {
          return getTxnProductLineValue(txn, field, 0);
        }
        return formatMergeDisplayValue(field, getAyalaRawFieldValue(record, field));
      });

      const sheetValues = values.map((value) => {
        const text = String(value ?? '').trim();
        if (!shouldSum) {
          return text;
        }
        if (/^-?\d+(?:\.\d+)?$/.test(text)) {
          return Number.parseFloat(text);
        }
        return text;
      });
      const sumValue = shouldSum
        ? formatAyalaAmount(values.reduce((acc, value) => acc + parseNumericAyala(value), 0))
        : '';

      const eodField = eodAlias[field] || field;
      const eodValue = String(eodCombinedValues.get(eodField) ?? '');

      const rowIndex = aoa.length;
      const displayFieldLabel = productRawFields.has(field) && maxProductLineCount > 1
        ? `${field} (1)`
        : field;
      const row = Array(eodFirstValueColIndex + 1).fill('');
      row[0] = displayFieldLabel;
      sheetValues.forEach((value, idx) => {
        row[firstTxnColIndex + idx] = value;
      });
      row[fieldCopyColIndex] = displayFieldLabel;
      if (hasMultiHourlyTerminal) {
        hourlyTerminalOrder.forEach((terNo, terIdx) => {
          const terTotal = shouldSum
            ? formatAyalaAmount(
              horizontalTxns.reduce((acc, txnKey) => {
                const txnTer = String(txnKey).split('|')[0] || '-';
                if (txnTer !== terNo) {
                  return acc;
                }
                const record = txnMap.get(txnKey);
                return acc + getAyalaNumericFieldValue(record || {}, field);
              }, 0)
            )
            : '';
          row[hourlyTerminalStartColIndex + terIdx] = terTotal;
        });
      }
      row[sumColIndex] = sumValue;
      row[separatorColIndex] = '';
      row[eodFieldColIndex] = eodField;
      row[eodFirstValueColIndex] = eodValue;
      aoa.push(row);
      fieldRowMap.set(field, rowIndex);
      if (productRawFields.has(field)) {
        productLineRowMap.set(`${field}#0`, rowIndex);
      }
      formulaRows.push(shouldSum);
      if (compareHighlightFields.has(field)) {
        comparisonRows.push({
          rowIndex,
          field,
          hourlyValue: values.reduce((acc, value) => acc + parseNumericAyala(value), 0),
          eodValue: parseNumericAyala(eodValue),
        });
      }
    });
    if (maxProductLineCount > 1) {
      for (let lineIdx = 1; lineIdx < maxProductLineCount; lineIdx += 1) {
        productFieldOrder.forEach((field) => {
          const values = horizontalTxns.map((txn) => {
            if (movedTxnKeys.has(txn)) {
              return '';
            }
            return getTxnProductLineValue(txn, field, lineIdx);
          });
          const displayFieldLabel = `${field} (${lineIdx + 1})`;
          const row = Array(eodFirstValueColIndex + 1).fill('');
          row[0] = displayFieldLabel;
          values.forEach((value, idx) => {
            row[firstTxnColIndex + idx] = String(value ?? '').trim();
          });
          row[fieldCopyColIndex] = displayFieldLabel;
          row[sumColIndex] = '';
          row[separatorColIndex] = '';
          row[eodFieldColIndex] = '';
          row[eodFirstValueColIndex] = '';
          aoa.push(row);
          productLineRowMap.set(`${field}#${lineIdx}`, aoa.length - 1);
          formulaRows.push(false);
        });
      }
    }

    const worksheet = XLSX.utils.aoa_to_sheet(aoa);

    const rawRows = Array.isArray(eodGrid.rows) ? eodGrid.rows : [];

    // Refund pair columns (right side of EOD), aligned to same field rows like manual checker.
    let refundMatrixMaxRow = -1;
    let refundMatrixEndCol = eodBlockEndColIndex;
    const pairSource = refundMatches.length > 0
      ? refundMatches
      : unmatchedRefunds.map((refund) => ({ refund, original: null }));
    if (hasRefundScenario && pairSource.length > 0) {
      pairSource.forEach((pair, pairIndex) => {
        const colOriginal = eodBlockEndColIndex + 1 + (pairIndex * 3);
        const colRefund = colOriginal + 1;
        const colAdjust = colRefund + 1;
        refundMatrixEndCol = Math.max(refundMatrixEndCol, colAdjust);

        const setByField = (field, col, value, numeric) => {
          if (!fieldRowMap.has(field)) {
            return;
          }
          const row = fieldRowMap.get(field);
          const addr = XLSX.utils.encode_cell({ r: row, c: col });
          refundMatrixMaxRow = Math.max(refundMatrixMaxRow, row);
          if (numeric) {
            worksheet[addr] = { t: 'n', v: roundAyalaAmount(parseNumericAyala(value)) };
          } else {
            worksheet[addr] = { t: 's', v: String(value ?? '') };
          }
        };

        const fillTxnColumn = (entry, col) => {
          orderedFields.forEach((field) => {
            if (productRawFields.has(field)) {
              const productValues = entry && entry.record
                ? getAyalaRawFieldValues(entry.record, field)
                  .map((value) => String(value ?? '').trim())
                  .map((value) => (value ? formatMergeDisplayValue(field, value) : ''))
                : [];
              for (let lineIdx = 0; lineIdx < maxProductLineCount; lineIdx += 1) {
                const row = productLineRowMap.get(`${field}#${lineIdx}`);
                if (row === undefined) {
                  continue;
                }
                const addr = XLSX.utils.encode_cell({ r: row, c: col });
                worksheet[addr] = { t: 's', v: String(productValues[lineIdx] || '') };
                refundMatrixMaxRow = Math.max(refundMatrixMaxRow, row);
              }
              return;
            }
            if (!fieldRowMap.has(field)) {
              return;
            }
            const row = fieldRowMap.get(field);
            const addr = XLSX.utils.encode_cell({ r: row, c: col });
            const raw = entry && entry.record ? getAyalaRawFieldValue(entry.record, field) : '';
            const display = formatMergeDisplayValue(field, raw);
            const text = String(display ?? '').trim();
            if (/^-?\d+(?:\.\d+)?$/.test(text)) {
              worksheet[addr] = { t: 'n', v: Number.parseFloat(text) };
            } else {
              worksheet[addr] = { t: 's', v: String(display ?? '') };
            }
            refundMatrixMaxRow = Math.max(refundMatrixMaxRow, row);
          });
        };

        const refund = pair.refund;
        const original = pair.original;

        fillTxnColumn(original || null, colOriginal);
        fillTxnColumn(refund, colRefund);
      });
    }

    // Keep product line rows visually consistent with numeric rows.
    const productRowIndexes = Array.from(new Set(Array.from(productLineRowMap.values())));
    productRowIndexes.forEach((rowIndex) => {
      for (let c = firstTxnColIndex; c <= refundMatrixEndCol; c += 1) {
        const addr = XLSX.utils.encode_cell({ r: rowIndex, c });
        if (!worksheet[addr]) {
          continue;
        }
        const baseStyle = worksheet[addr].s || {};
        const baseAlign = baseStyle.alignment || {};
        worksheet[addr].s = {
          ...baseStyle,
          alignment: {
            ...baseAlign,
            horizontal: 'right',
            vertical: 'center',
          },
        };
      }
    });

    // Write SUM formulas so users can inspect exact calculation in Excel.
    for (let r = 0; r < formulaRows.length; r += 1) {
      if (!formulaRows[r]) {
        continue;
      }
      const rowNumber = r + 1; // Excel is 1-based
      const startCol = XLSX.utils.encode_col(firstTxnColIndex);
      const endCol = XLSX.utils.encode_col(lastTxnColIndex);
      const sumAddress = XLSX.utils.encode_cell({ r, c: sumColIndex });
      const rowValues = aoa[r] || [];
      const computed = rowValues
        .slice(firstTxnColIndex, lastTxnColIndex + 1)
        .reduce((acc, value) => acc + parseNumericAyala(value), 0);
      const fieldName = String(rowValues[0] || '').trim();
      let sumFormula = `SUM(${startCol}${rowNumber}:${endCol}${rowNumber})`;
      let computedValue = roundAyalaAmount(computed);
      const inAdjustmentRange = adjustmentFields.includes(fieldName);
      if (inAdjustmentRange && fieldName === 'GROSS_SLS' && movedOriginalGrossTotal !== 0) {
        const movedOriginalLiteral = movedOriginalGrossTotal.toFixed(2);
        sumFormula += `+${movedOriginalLiteral}`;
        computedValue = roundAyalaAmount(computedValue + movedOriginalGrossTotal);
      }
      if (fieldName === 'REFUND_AMT' && movedRefundTotal > 0) {
        const movedRefundLiteral = movedRefundTotal.toFixed(2);
        sumFormula += `+${movedRefundLiteral}`;
        computedValue = roundAyalaAmount(computedValue + movedRefundTotal);
      }
      if (paymentAdjustmentFields.includes(fieldName)) {
        const paymentAdjust = roundAyalaAmount(movedPaymentAdjustByField[fieldName] || 0);
        if (!isSameAyalaAmount(paymentAdjust, 0)) {
          const paymentAdjustLiteral = paymentAdjust.toFixed(2);
          sumFormula += paymentAdjust >= 0 ? `+${paymentAdjustLiteral}` : paymentAdjustLiteral;
          computedValue = roundAyalaAmount(computedValue + paymentAdjust);
        }
      }
      worksheet[sumAddress] = {
        t: 'n',
        f: sumFormula,
        v: computedValue,
      };

      if (hasMultiHourlyTerminal) {
        hourlyTerminalOrder.forEach((terNo, terIdx) => {
          const range = terminalRangeMap.get(terNo);
          if (!range) {
            return;
          }
          const startTerCol = XLSX.utils.encode_col(range.start);
          const endTerCol = XLSX.utils.encode_col(range.end);
          const terAddress = XLSX.utils.encode_cell({ r, c: hourlyTerminalStartColIndex + terIdx });
          const terComputed = rowValues
            .slice(range.start, range.end + 1)
            .reduce((acc, value) => acc + parseNumericAyala(value), 0);
          worksheet[terAddress] = {
            t: 'n',
            f: `SUM(${startTerCol}${rowNumber}:${endTerCol}${rowNumber})`,
            v: roundAyalaAmount(terComputed),
          };
        });
      }
    }

    // Try to visually separate hourly totals and EOD with a dark separator column.
    for (let r = 0; r < aoa.length; r += 1) {
      const addr = XLSX.utils.encode_cell({ r, c: separatorColIndex });
      worksheet[addr] = {
        t: 's',
        v: '',
        s: {
          fill: {
            patternType: 'solid',
            fgColor: { rgb: '000000' },
            bgColor: { rgb: '000000' },
          },
        },
      };
    }

    // Paste EOD list, by terminal columns for multi-terminal EOD.
    const rawFieldRowMap = new Map();
    if (rawRows.length > 0) {
      if (eodGrid.isMulti) {
        const headerRowIndex = 1; // row 2
        const headerFieldAddr = XLSX.utils.encode_cell({ r: headerRowIndex, c: eodFieldColIndex });
        worksheet[headerFieldAddr] = { t: 's', v: 'TER_NO' };
        eodGrid.terminals.forEach((terNo, tIdx) => {
          const headerAddr = XLSX.utils.encode_cell({ r: headerRowIndex, c: eodFirstValueColIndex + tIdx });
          worksheet[headerAddr] = { t: 's', v: String(terNo || '-') };
        });
        if (eodTotalColIndex >= 0) {
          const totalHeaderAddr = XLSX.utils.encode_cell({ r: headerRowIndex, c: eodTotalColIndex });
          worksheet[totalHeaderAddr] = { t: 's', v: 'TOTAL' };
        }
      }

      rawRows.forEach((row, idx) => {
        const rowIndex = 2 + idx; // row 3 in Excel is index 2
        const keyAddr = XLSX.utils.encode_cell({ r: rowIndex, c: eodFieldColIndex });
        rawFieldRowMap.set(String(row.field ?? '').trim(), rowIndex);
        worksheet[keyAddr] = { t: 's', v: String(row.field ?? '') };

        if (eodGrid.isMulti) {
          const values = Array.isArray(row.values) ? row.values : [];
          eodGrid.terminals.forEach((_, tIdx) => {
            const colIndex = eodFirstValueColIndex + tIdx;
            const valueAddr = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
            worksheet[valueAddr] = { t: 's', v: String(values[tIdx] ?? '') };
          });
          if (eodTotalColIndex >= 0) {
            const totalAddr = XLSX.utils.encode_cell({ r: rowIndex, c: eodTotalColIndex });
            worksheet[totalAddr] = { t: 's', v: String(row.total ?? '') };
          }
        } else {
          const valAddr = XLSX.utils.encode_cell({ r: rowIndex, c: eodFirstValueColIndex });
          const singleValue = Array.isArray(row.values) && row.values.length > 0
            ? row.values[0]
            : '';
          worksheet[valAddr] = { t: 's', v: String(singleValue ?? '') };
        }
      });

      // Expand worksheet range so all manually written EOD rows are visible in Excel.
      const existingRange = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
      const lastRawRowIndex = 2 + rawRows.length - 1;
      existingRange.e.r = Math.max(existingRange.e.r, lastRawRowIndex);
      const rawEndCol = eodGrid.isMulti && eodTotalColIndex >= 0
        ? eodTotalColIndex
        : eodFirstValueColIndex;
      existingRange.e.c = Math.max(existingRange.e.c, rawEndCol);
      if (refundMatrixMaxRow >= 0) {
        existingRange.e.r = Math.max(existingRange.e.r, refundMatrixMaxRow);
        existingRange.e.c = Math.max(existingRange.e.c, refundMatrixEndCol);
      }
      worksheet['!ref'] = XLSX.utils.encode_range(existingRange);
    }
    if (rawRows.length === 0 && refundMatrixMaxRow >= 0) {
      const existingRange = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
      existingRange.e.r = Math.max(existingRange.e.r, refundMatrixMaxRow);
      existingRange.e.c = Math.max(existingRange.e.c, refundMatrixEndCol);
      worksheet['!ref'] = XLSX.utils.encode_range(existingRange);
    }

    const matchStyle = {
      fill: {
        patternType: 'solid',
        fgColor: { rgb: 'C6EFCE' },
      },
      font: { color: { rgb: '006100' }, bold: true },
    };
    const mismatchStyle = {
      fill: {
        patternType: 'solid',
        fgColor: { rgb: 'FFC7CE' },
      },
      font: { color: { rgb: '9C0006' }, bold: true },
    };
    const eodTotalCompareFields = new Set(['GROSS_SLS', 'VAT_AMNT', 'VATABLE_SLS', 'REFUND_AMT']);
    comparisonRows.forEach((item) => {
      const hourlyAddr = XLSX.utils.encode_cell({ r: item.rowIndex, c: sumColIndex });
      const computedHourlyValue = worksheet[hourlyAddr]
        ? parseNumericAyala(worksheet[hourlyAddr].v)
        : item.hourlyValue;
      const style = isSameAyalaAmount(computedHourlyValue, item.eodValue) ? matchStyle : mismatchStyle;
      if (worksheet[hourlyAddr]) {
        worksheet[hourlyAddr].s = style;
      }

      let eodRowIndex = item.rowIndex;
      if (rawFieldRowMap.size > 0 && rawFieldRowMap.has(item.field)) {
        eodRowIndex = rawFieldRowMap.get(item.field);
      }
      const useTotalCol = eodGrid.isMulti && eodTotalColIndex >= 0 && eodTotalCompareFields.has(item.field);
      if (eodGrid.isMulti && !useTotalCol) {
        eodRowIndex = item.rowIndex;
      }
      const eodValueCol = useTotalCol ? eodTotalColIndex : eodFirstValueColIndex;
      const eodAddr = XLSX.utils.encode_cell({ r: eodRowIndex, c: eodValueCol });
      if (worksheet[eodAddr]) {
        worksheet[eodAddr].s = style;
      }
    });

    const totalColumns = (
      eodGrid.isMulti && eodTotalColIndex >= 0
        ? eodTotalColIndex + 1
        : eodFirstValueColIndex + 1
    );
    const finalTotalColumns = Math.max(totalColumns, refundMatrixEndCol + 1);
    worksheet['!cols'] = Array.from({ length: finalTotalColumns }, (_, index) => {
      if (index === 0) {
        return { wch: 18 };
      }
      if (index === fieldCopyColIndex) {
        return { wch: 18 };
      }
      if (hasMultiHourlyTerminal && index >= hourlyTerminalStartColIndex && index < sumColIndex) {
        return { wch: 12 };
      }
      if (index === sumColIndex) {
        return { wch: 14 };
      }
      if (index === separatorColIndex) {
        return { wch: 1.5 };
      }
      if (index === eodFieldColIndex) {
        return { wch: 18 };
      }
      if (index >= eodFirstValueColIndex) {
        return { wch: 14 };
      }
      if (index > eodBlockEndColIndex && index <= refundMatrixEndCol) {
        return { wch: 14 };
      }
      return { wch: 11 };
    });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Checker');
    const txnValidationSheet = buildAyalaTransactionFormulaSheet(hourlyRecords);
    XLSX.utils.book_append_sheet(workbook, txnValidationSheet, 'Txn Validation');
    return workbook;
  }

  function buildAyalaTransactionFormulaSheet(hourlyRecords) {
    const grossTotalFields = [
      'VAT_AMNT', 'VATABLE_SLS', 'NONVAT_SLS', 'VATEXEMPT_SLS', 'VATEXEMPT_AMNT',
      'LOCAL_TAX', 'PWD_DISC', 'SNRCIT_DISC', 'EMPLO_DISC', 'AYALA_DISC', 'STORE_DISC',
      'OTHER_DISC', 'SCHRGE_AMT', 'OTHER_SCHR',
    ];
    const paymentTotalFields = [
      'CASH_SLS', 'OTHERSL_SLS', 'CHECK_SLS', 'GC_SLS', 'MASTERCARD_SLS', 'VISA_SLS', 'AMEX_SLS',
      'DINERS_SLS', 'JCB_SLS', 'GCASH_SLS', 'PAYMAYA_SLS', 'ALIPAY_SLS', 'WECHAT_SLS', 'GRAB_SLS',
      'FOODPANDA_SLS', 'MASTERDEBIT_SLS', 'VISADEBIT_SLS', 'PAYPAL_SLS', 'ONLINE_SLS', 'OPEN_SALES',
      'OPEN_SALES_2', 'OPEN_SALES_3', 'OPEN_SALES_4', 'OPEN_SALES_5', 'OPEN_SALES_6', 'OPEN_SALES_7',
      'OPEN_SALES_8', 'OPEN_SALES_9', 'OPEN_SALES_10', 'OPEN_SALES_11', 'GC_EXCESS',
    ];
    const paymentBaseFields = [
      'VAT_AMNT', 'VATABLE_SLS', 'NONVAT_SLS', 'VATEXEMPT_SLS', 'SCHRGE_AMT', 'OTHER_SCHR', 'LOCAL_TAX',
    ];

    const headers = [
      'TRANSACTION_NO',
      'TER_NO',
      'GROSS_SLS',
      ...grossTotalFields,
      ...paymentTotalFields,
      'GROSS_TOTAL',
      'DISCREPANCY',
      'GROSS_DISCREPANCY',
      'PAYMENT',
      'PAYMENT_TOTAL',
      'PAYMENT_DISCREPANCY',
      'HAS_ERROR',
    ];

    const sorted = (hourlyRecords || [])
      .map((record) => ({
        txn: String(record.TRANSACTION_NO || '').trim(),
        ter: String(record.TER_NO || '').trim(),
        key: getAyalaTransactionCompositeKey(record),
        record,
      }))
      .filter((item) => item.txn && item.key && item.key !== '|')
      .sort((a, b) => {
        return compareAyalaTxnEntries(a.ter, a.txn, b.ter, b.txn);
      });

    const aoa = [headers];
    const rowRecords = [null];
    const seen = new Set();
    sorted.forEach((item) => {
      if (seen.has(item.key)) {
        return;
      }
      seen.add(item.key);
      const row = [
        item.txn,
        String(item.record.TER_NO || ''),
        roundAyalaAmount(parseNumericAyala(item.record.GROSS_SLS)),
        ...grossTotalFields.map((field) => roundAyalaAmount(parseNumericAyala(item.record[field]))),
        ...paymentTotalFields.map((field) => roundAyalaAmount(getAyalaNumericFieldValue(item.record, field))),
        0,
        0,
        0,
        0,
        0,
        0,
        '',
      ];
      aoa.push(row);
      rowRecords.push(item.record);
    });

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const matchStyle = {
      fill: {
        patternType: 'solid',
        fgColor: { rgb: 'C6EFCE' },
      },
      font: { color: { rgb: '006100' }, bold: true },
    };
    const mismatchStyle = {
      fill: {
        patternType: 'solid',
        fgColor: { rgb: 'FFC7CE' },
      },
      font: { color: { rgb: '9C0006' }, bold: true },
    };
    const grossCol = 2; // C
    const grossSourceStartCol = grossCol + 1;
    const grossSourceEndCol = grossSourceStartCol + grossTotalFields.length - 1;
    const paymentSourceStartCol = grossSourceEndCol + 1;
    const paymentSourceEndCol = paymentSourceStartCol + paymentTotalFields.length - 1;
    const grossTotalCol = paymentSourceEndCol + 1;
    const legacyDiscrepancyCol = grossTotalCol + 1;
    const grossDiscrepancyCol = legacyDiscrepancyCol + 1;
    const paymentBaseCol = grossDiscrepancyCol + 1;
    const paymentTotalCol = paymentBaseCol + 1;
    const paymentDiscrepancyCol = paymentTotalCol + 1;
    const hasErrorCol = paymentDiscrepancyCol + 1;
    const headerIndexMap = new Map(headers.map((header, idx) => [header, idx]));

    for (let r = 1; r < aoa.length; r += 1) {
      const excelRow = r + 1;
      const grossAddr = XLSX.utils.encode_cell({ r, c: grossCol });
      const grossTotalAddr = XLSX.utils.encode_cell({ r, c: grossTotalCol });
      const legacyDiscrepancyAddr = XLSX.utils.encode_cell({ r, c: legacyDiscrepancyCol });
      const grossDiscrepancyAddr = XLSX.utils.encode_cell({ r, c: grossDiscrepancyCol });
      const paymentBaseAddr = XLSX.utils.encode_cell({ r, c: paymentBaseCol });
      const paymentTotalAddr = XLSX.utils.encode_cell({ r, c: paymentTotalCol });
      const paymentDiscrepancyAddr = XLSX.utils.encode_cell({ r, c: paymentDiscrepancyCol });
      const hasErrorAddr = XLSX.utils.encode_cell({ r, c: hasErrorCol });

      const grossSourceStart = XLSX.utils.encode_col(grossSourceStartCol);
      const grossSourceEnd = XLSX.utils.encode_col(grossSourceEndCol);
      const paymentSourceStart = XLSX.utils.encode_col(paymentSourceStartCol);
      const paymentSourceEnd = XLSX.utils.encode_col(paymentSourceEndCol);
      const grossValue = aoa[r][grossCol];
      const grossTotalValue = aoa[r]
        .slice(grossSourceStartCol, grossSourceEndCol + 1)
        .reduce((sum, v) => sum + parseNumericAyala(v), 0);
      const rowRecord = rowRecords[r] || {};
      const paymentTotalValueRaw = paymentTotalFields
        .reduce((sum, field) => sum + getAyalaNumericFieldValue(rowRecord, field), 0);
      const gcExcessValue = getAyalaNumericFieldValue(rowRecord, 'GC_EXCESS');
      const paymentTotalValue = paymentTotalValueRaw - (gcExcessValue * 2);
      const paymentBaseValue = paymentBaseFields
        .reduce((sum, field) => sum + getAyalaNumericFieldValue(rowRecord, field), 0);
      const grossDiscrepancyValue = getAyalaFlexibleDiscrepancy(grossValue, grossTotalValue, { allowOppositeSignSameAbs: true });
      const paymentDiscrepancyValue = getAyalaPaymentDiscrepancy(paymentBaseValue, paymentTotalValue);
      const grossColRef = `${XLSX.utils.encode_col(grossCol)}${excelRow}`;
      const grossTotalColRef = `${XLSX.utils.encode_col(grossTotalCol)}${excelRow}`;
      const paymentBaseColRef = `${XLSX.utils.encode_col(paymentBaseCol)}${excelRow}`;
      const paymentTotalColRef = `${XLSX.utils.encode_col(paymentTotalCol)}${excelRow}`;
      const gcExcessColRef = `${XLSX.utils.encode_col(headerIndexMap.get('GC_EXCESS'))}${excelRow}`;

      ws[grossTotalAddr] = {
        t: 'n',
        f: `SUM(${grossSourceStart}${excelRow}:${grossSourceEnd}${excelRow})`,
        v: roundAyalaAmount(grossTotalValue),
      };
      ws[grossDiscrepancyAddr] = {
        t: 'n',
        f: `IF(AND(SIGN(${grossColRef})<>0,SIGN(${grossTotalColRef})<>0,SIGN(${grossColRef})<>SIGN(${grossTotalColRef}),ABS(ROUND(ABS(${grossColRef}),2)-ROUND(ABS(${grossTotalColRef}),2))<=${AYALA_DISCREPANCY_TOLERANCE}),0,ABS(${grossColRef}-${grossTotalColRef}))`,
        v: roundAyalaAmount(grossDiscrepancyValue),
      };
      ws[legacyDiscrepancyAddr] = {
        t: 'n',
        f: `IF(AND(SIGN(${grossColRef})<>0,SIGN(${grossTotalColRef})<>0,SIGN(${grossColRef})<>SIGN(${grossTotalColRef}),ABS(ROUND(ABS(${grossColRef}),2)-ROUND(ABS(${grossTotalColRef}),2))<=${AYALA_DISCREPANCY_TOLERANCE}),0,ABS(${grossColRef}-${grossTotalColRef}))`,
        v: roundAyalaAmount(grossDiscrepancyValue),
      };
      ws[paymentTotalAddr] = {
        t: 'n',
        f: `SUM(${paymentSourceStart}${excelRow}:${paymentSourceEnd}${excelRow})-(${gcExcessColRef}*2)`,
        v: roundAyalaAmount(paymentTotalValue),
      };
      ws[paymentBaseAddr] = {
        t: 'n',
        f: `SUM(${paymentBaseFields.map((field) => {
          const idx = headerIndexMap.get(field);
          return `${XLSX.utils.encode_col(idx)}${excelRow}`;
        }).join(',')})`,
        v: roundAyalaAmount(paymentBaseValue),
      };
      ws[paymentDiscrepancyAddr] = {
        t: 'n',
        f: `IF(AND(SIGN(${paymentBaseColRef})<>0,SIGN(${paymentTotalColRef})<>0,SIGN(${paymentBaseColRef})<>SIGN(${paymentTotalColRef}),ABS(ROUND(ABS(${paymentBaseColRef}),2)-ROUND(ABS(${paymentTotalColRef}),2))<=${AYALA_DISCREPANCY_TOLERANCE}),0,ABS(${paymentBaseColRef}-${paymentTotalColRef}))`,
        v: roundAyalaAmount(paymentDiscrepancyValue),
      };
      ws[hasErrorAddr] = {
        t: 's',
        f: `IF(OR(${XLSX.utils.encode_col(grossDiscrepancyCol)}${excelRow}>${AYALA_DISCREPANCY_TOLERANCE},${XLSX.utils.encode_col(paymentDiscrepancyCol)}${excelRow}>${AYALA_DISCREPANCY_TOLERANCE}),"ERROR","OK")`,
        v: (grossDiscrepancyValue > AYALA_DISCREPANCY_TOLERANCE || paymentDiscrepancyValue > AYALA_DISCREPANCY_TOLERANCE) ? 'ERROR' : 'OK',
      };
      const rowStyle = (grossDiscrepancyValue > AYALA_DISCREPANCY_TOLERANCE || paymentDiscrepancyValue > AYALA_DISCREPANCY_TOLERANCE)
        ? mismatchStyle
        : matchStyle;
      if (ws[grossAddr]) ws[grossAddr].s = rowStyle;
      if (ws[grossTotalAddr]) ws[grossTotalAddr].s = rowStyle;
      if (ws[legacyDiscrepancyAddr]) ws[legacyDiscrepancyAddr].s = rowStyle;
      if (ws[grossDiscrepancyAddr]) ws[grossDiscrepancyAddr].s = rowStyle;
      if (ws[paymentBaseAddr]) ws[paymentBaseAddr].s = rowStyle;
      if (ws[paymentTotalAddr]) ws[paymentTotalAddr].s = rowStyle;
      if (ws[paymentDiscrepancyAddr]) ws[paymentDiscrepancyAddr].s = rowStyle;
      if (ws[hasErrorAddr]) ws[hasErrorAddr].s = rowStyle;
    }

    ws['!cols'] = headers.map((header) => {
      if (header === 'HAS_ERROR') {
        return { wch: 12 };
      }
      if (header === 'TRANSACTION_NO' || header === 'TER_NO') {
        return { wch: 14 };
      }
      if (header === 'GROSS_SLS' || header === 'GROSS_TOTAL' || header === 'DISCREPANCY' || header === 'GROSS_DISCREPANCY' || header === 'PAYMENT' || header === 'PAYMENT_TOTAL' || header === 'PAYMENT_DISCREPANCY') {
        return { wch: 14 };
      }
      return { wch: 12 };
    });
    return ws;
  }

  function parseOrderedReceiptNumbersFromText(textContent) {
    const lines = textContent.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
    if (lines.length === 0) {
      return [];
    }

    const normalize = (value) => String(value || '').toLowerCase().replace(/\s+/g, '').trim();

    // Find a real table header row containing Receipt# in first lines.
    let headerRowIndex = -1;
    let receiptColumnIndex = -1;
    let delimiter = null;

    for (let r = 0; r < Math.min(lines.length, 50); r += 1) {
      const rowDelimiter = detectCsvDelimiter(lines[r]);
      if (!rowDelimiter) {
        continue;
      }

      const cells = lines[r].split(rowDelimiter).map((cell) => normalize(cell));
      const idx = cells.findIndex((cell) => cell === 'receipt#' || cell === 'reciept#');
      if (idx >= 0) {
        headerRowIndex = r;
        receiptColumnIndex = idx;
        delimiter = rowDelimiter;
        break;
      }
    }

    if (headerRowIndex < 0 || receiptColumnIndex < 0 || !delimiter) {
      return [];
    }

    let grossColumnIndex = -1;
    let customerColumnIndex = -1;
    let dateColumnIndex = -1;
    const headerCells = lines[headerRowIndex]
      .split(delimiter)
      .map((cell) => String(cell || '').trim());
    grossColumnIndex = headerCells.findIndex((cell) => isGrossHeaderCell(cell));
    customerColumnIndex = headerCells.findIndex((cell) => isCustomerHeaderCell(cell));
    dateColumnIndex = headerCells.findIndex((cell) => isDateHeaderCell(cell));

    const values = [];
    for (let i = headerRowIndex + 1; i < lines.length; i += 1) {
      const cells = lines[i].split(delimiter);
      const raw = (cells[receiptColumnIndex] || '').trim();
      const match = raw.match(/\d+/);
      if (match) {
        values.push({
          receipt: Number.parseInt(match[0], 10),
          row: i + 1,
          grossRaw: grossColumnIndex >= 0 ? String(cells[grossColumnIndex] ?? '').trim() : '',
          customerRaw: customerColumnIndex >= 0 ? String(cells[customerColumnIndex] ?? '').trim() : '',
          dateRaw: dateColumnIndex >= 0 ? String(cells[dateColumnIndex] ?? '').trim() : '',
        });
      }
    }

    return values.filter((value) => Number.isFinite(value.receipt));
  }

  function extractReceiptNumbersFromSheetRows(rows) {
    if (!rows || rows.length === 0) {
      return [];
    }

    const normalize = (value) => String(value || '').toLowerCase().replace(/\s+/g, '').trim();

    let headerRowIndex = -1;
    let receiptColumnIndex = -1;

    for (let r = 0; r < Math.min(rows.length, 80); r += 1) {
      const row = rows[r] || [];
      const idx = row.findIndex((cell) => {
        const text = normalize(cell);
        return text === 'receipt#' || text === 'reciept#';
      });

      if (idx >= 0) {
        headerRowIndex = r;
        receiptColumnIndex = idx;
        break;
      }
    }

    if (headerRowIndex < 0 || receiptColumnIndex < 0) {
      return [];
    }

    let grossColumnIndex = -1;
    let customerColumnIndex = -1;
    let dateColumnIndex = -1;
    const headerRow = rows[headerRowIndex] || [];
    grossColumnIndex = headerRow.findIndex((cell) => isGrossHeaderCell(cell));
    customerColumnIndex = headerRow.findIndex((cell) => isCustomerHeaderCell(cell));
    dateColumnIndex = headerRow.findIndex((cell) => isDateHeaderCell(cell));

    const values = [];
    for (let i = headerRowIndex + 1; i < rows.length; i += 1) {
      const row = rows[i] || [];
      const raw = String(row[receiptColumnIndex] ?? '').trim();
      const match = raw.match(/\d+/);
      if (match) {
        values.push({
          receipt: Number.parseInt(match[0], 10),
          row: i + 1,
          grossRaw: grossColumnIndex >= 0 ? String(row[grossColumnIndex] ?? '').trim() : '',
          customerRaw: customerColumnIndex >= 0 ? String(row[customerColumnIndex] ?? '').trim() : '',
          dateRaw: dateColumnIndex >= 0 ? String(row[dateColumnIndex] ?? '').trim() : '',
        });
      }
    }

    return values.filter((value) => Number.isFinite(value.receipt));
  }

  function showAnalysisResult(orderedNumbers, missing, issues, duplicates, rollbacks, duplicateDetailMismatches) {
    if (!missingSummary || !detectedCount || !missingCount || !outOfOrderCount || !missingListOutput || !sqlActions) {
      return;
    }

    missingOrNumbers = missing;

    detectedCount.textContent = String(orderedNumbers.length);
    missingCount.textContent = String(missing.length);
    if (duplicateCount) {
      duplicateCount.textContent = String((duplicates || []).length);
    }
    if (duplicateDetailMismatchCount) {
      duplicateDetailMismatchCount.textContent = String((duplicateDetailMismatches || []).length);
    }
    if (rollbackCount) {
      rollbackCount.textContent = String((rollbacks || []).length);
    }
    outOfOrderCount.textContent = String(issues.length);

    const lines = [];
    lines.push('Missing OR#:');
    lines.push(missing.length > 0 ? missing.join(', ') : 'None');
    lines.push('');
    lines.push('Duplicate OR# rows:');
    lines.push((duplicates && duplicates.length > 0)
      ? duplicates.map((item) => `Row ${item.row}: ${item.value}`).join('\n')
      : 'None');
    lines.push('');
    lines.push('Roll back OR# rows:');
    lines.push((rollbacks && rollbacks.length > 0)
      ? rollbacks.map((item) => String(item.issue || `Row ${item.row}: ${item.from} -> ${item.to}`)).join('\n')
      : 'None');

    missingListOutput.value = lines.join('\n');

    missingSummary.hidden = false;
    sqlActions.hidden = false;
    if (analyzeStatus) {
      analyzeStatus.hidden = false;
      analyzeStatus.textContent = 'Analysis complete.';
    }
  }

  function resetMissingOutput() {
    missingOrNumbers = [];
    if (detectedCount) detectedCount.textContent = '0';
    if (missingCount) missingCount.textContent = '0';
    if (duplicateCount) duplicateCount.textContent = '0';
    if (duplicateDetailMismatchCount) duplicateDetailMismatchCount.textContent = '0';
    if (rollbackCount) rollbackCount.textContent = '0';
    if (outOfOrderCount) outOfOrderCount.textContent = '0';

    if (missingSummary) {
      missingSummary.hidden = true;
    }
    if (sqlActions) {
      sqlActions.hidden = true;
    }
    if (sqlResultOutput) {
      sqlResultOutput.hidden = true;
      sqlResultOutput.value = '';
    }
    if (copyMissingSqlBtn) {
      copyMissingSqlBtn.disabled = true;
    }
    if (copyMissingSqlRow) {
      copyMissingSqlRow.hidden = true;
    }
    if (missingListOutput) {
      missingListOutput.value = '';
    }
    if (analyzeStatus) {
      analyzeStatus.hidden = true;
      analyzeStatus.textContent = 'Analyzing file, please wait...';
    }
  }

  function processUploadedText(textContent) {
    const orderedNumbers = parseOrderedReceiptNumbersFromText(textContent);
    const { missing, issues, duplicates, rollbacks, duplicateDetailMismatches } = analyzeSequence(orderedNumbers);
    showAnalysisResult(orderedNumbers, missing, issues, duplicates, rollbacks, duplicateDetailMismatches);

    if (orderedNumbers.length === 0 && missingListOutput) {
      missingListOutput.value = 'No Receipt# column found in uploaded text/CSV.';
      if (analyzeStatus) {
        analyzeStatus.hidden = false;
        analyzeStatus.textContent = 'Analysis complete: Receipt# column not found.';
      }
    }
  }

  function processUploadedWorkbook(arrayBuffer) {
    if (typeof XLSX === 'undefined') {
      if (missingSummary && detectedCount && missingCount && outOfOrderCount && missingListOutput) {
        detectedCount.textContent = '0';
        missingCount.textContent = '0';
        if (duplicateCount) duplicateCount.textContent = '0';
        if (duplicateDetailMismatchCount) duplicateDetailMismatchCount.textContent = '0';
        if (rollbackCount) rollbackCount.textContent = '0';
        outOfOrderCount.textContent = '0';
        missingListOutput.value = 'Excel parser failed to load.';
        missingSummary.hidden = false;
      }
      if (analyzeStatus) {
        analyzeStatus.hidden = false;
        analyzeStatus.textContent = 'Analysis failed: Excel parser not loaded.';
      }
      return;
    }

    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
    const orderedNumbers = extractReceiptNumbersFromSheetRows(rows);
    const { missing, issues, duplicates, rollbacks, duplicateDetailMismatches } = analyzeSequence(orderedNumbers);

    showAnalysisResult(orderedNumbers, missing, issues, duplicates, rollbacks, duplicateDetailMismatches);

    if (orderedNumbers.length === 0 && missingListOutput) {
      missingListOutput.value = 'No Receipt# column found in the worksheet.';
      if (analyzeStatus) {
        analyzeStatus.hidden = false;
        analyzeStatus.textContent = 'Analysis complete: Receipt# column not found.';
      }
    }
  }

  function buildSelectSqlFromMissing() {
    if (missingOrNumbers.length === 0) {
      return '-- No missing OR numbers found.';
    }

    return [
      'SELECT *',
      'FROM pos_sale',
      'WHERE fdocument_no IN (',
      `${formatSqlNumberList(missingOrNumbers)}`,
      ');',
    ].join('\n');
  }

  function buildDeleteSqlFromMissing() {
    if (missingOrNumbers.length === 0) {
      return '-- No missing OR numbers found.';
    }

    return [
      'DELETE FROM pos_sale ',
      'WHERE fdocument_no NOT IN (',
      `${formatSqlNumberList(missingOrNumbers)}`,
      ');',
      '',
      'DELETE FROM pos_sale_payment ',
      'WHERE frecno NOT IN (SELECT frecno FROM pos_sale);',
      '',
      'DELETE FROM pos_sale_product ',
      'WHERE frecno NOT IN (SELECT frecno FROM pos_sale);',
    ].join('\n');
  }

  function parseFrecnoList(rawText) {
    const matches = String(rawText || '').match(/\d+/g);
    if (!matches) {
      return [];
    }
    return matches.map((value) => value.trim()).filter(Boolean);
  }

  function findDuplicateFrecnos(values) {
    const counts = new Map();
    values.forEach((value) => {
      counts.set(value, (counts.get(value) || 0) + 1);
    });

    return Array.from(counts.entries())
      .filter(([, count]) => count > 1)
      .map(([value]) => value);
  }

  function formatFrecnoSqlList(values) {
    return values.map((value) => `'${String(value)}'`).join(', ');
  }

  function extractUniqueFrecnos(rawText) {
    const parsed = parseFrecnoList(rawText);
    if (parsed.length === 0) {
      return [];
    }
    const seen = new Set();
    const unique = [];
    parsed.forEach((value) => {
      if (seen.has(value)) {
        return;
      }
      seen.add(value);
      unique.push(value);
    });
    return unique;
  }

  function buildSelectSqlFromFrecnoList(rawText) {
    const values = extractUniqueFrecnos(rawText);
    if (values.length === 0) {
      return '-- Please enter at least one frecno value.';
    }
    const listSql = formatFrecnoSqlList(values);
    return [
      `SELECT * FROM pos_sale WHERE frecno IN (${listSql});`,
      '',
      `SELECT * FROM pos_sale_payment WHERE frecno IN (${listSql});`,
      '',
      `SELECT * FROM pos_sale_product WHERE frecno IN (${listSql});`,
    ].join('\n');
  }

  function buildDeleteSqlFromFrecnoList(rawText) {
    const values = extractUniqueFrecnos(rawText);
    if (values.length === 0) {
      return '-- Please enter at least one frecno value.';
    }
    const listSql = formatFrecnoSqlList(values);
    return [
      `DELETE from pos_sale where frecno not in (${listSql});`,
      '',
      `DELETE from pos_sale_payment where frecno not in (${listSql});`,
      '',
      `DELETE from pos_sale_product where frecno not in (${listSql});`,
    ].join('\n');
  }

  function normalizeFrecnoInputToVertical() {
    if (!frecnoListSqlInput) {
      return;
    }
    const values = extractUniqueFrecnos(frecnoListSqlInput.value);
    if (values.length === 0) {
      return;
    }
    frecnoListSqlInput.value = values.join('\n');
  }

  function isIntegerText(value) {
    return /^\d+$/.test(String(value || '').trim());
  }

  async function buildFrecnoUpdateSql(oldFrecnos, latestStart) {
    const tables = ['pos_sale', 'pos_sale_payment', 'pos_sale_product'];
    const whenLines = [];
    const yieldEvery = 500;

    for (let index = 0; index < oldFrecnos.length; index += 1) {
      const oldFrecno = oldFrecnos[index];
      const nextFrecno = latestStart + index;
      whenLines.push(`    WHEN "${oldFrecno}" THEN "${nextFrecno}"`);

      if (index > 0 && index % yieldEvery === 0) {
        if (frecnoProgress) {
          frecnoProgress.hidden = false;
          frecnoProgress.textContent = `Finalizing SQL... ${index}/${oldFrecnos.length}`;
        }
        await new Promise((resolve) => window.setTimeout(resolve, 0));
      }
    }

    const blocks = tables.map((table) => [
      `UPDATE ${table} `,
      'SET frecno = ',
      '  CASE frecno',
      ...whenLines,
      '',
      '    ELSE frecno',
      '  END;',
    ].join('\n'));

    return blocks.join('\n\n');
  }

  function openMissingOrModal() {
    if (!missingOrModal) {
      return;
    }
    missingOrModal.hidden = false;
    document.body.style.overflow = 'hidden';
    if (orFileInput) {
      orFileInput.focus();
    }
  }

  function closeMissingOrModalFn() {
    if (!missingOrModal) {
      return;
    }
    missingOrModal.hidden = true;
    document.body.style.overflow = '';
  }

  function openFrecnoModal() {
    if (!frecnoModal) {
      return;
    }
    frecnoModal.hidden = false;
    document.body.style.overflow = 'hidden';
    if (oldFrecnoInput) {
      oldFrecnoInput.focus();
    }
    if (!frecnoOutput || !frecnoOutput.value.trim()) {
      hideFrecnoCopyButton();
    }
  }

  function closeFrecnoModalFn() {
    if (!frecnoModal) {
      return;
    }
    frecnoModal.hidden = true;
    document.body.style.overflow = '';
  }

  function openFrecnoListSqlModal() {
    if (!frecnoListSqlModal) {
      return;
    }
    frecnoListSqlModal.hidden = false;
    document.body.style.overflow = 'hidden';
    if (frecnoListSqlInput) {
      frecnoListSqlInput.focus();
    }
  }

  function closeFrecnoListSqlModalFn() {
    if (!frecnoListSqlModal) {
      return;
    }
    frecnoListSqlModal.hidden = true;
    document.body.style.overflow = '';
  }

  function openXmlRenameModal() {
    if (!xmlRenameModal) {
      return;
    }
    xmlRenameModal.hidden = false;
    document.body.style.overflow = 'hidden';
    if (xmlRenameFileInput) {
      xmlRenameFileInput.focus();
    }
    if (xmlRenameStatus) {
      xmlRenameStatus.hidden = false;
      xmlRenameStatus.textContent = 'Ready.';
    }
  }

  function openReportDocxModal() {
    if (!reportDocxModal) {
      return;
    }
    reportDocxModal.hidden = false;
    document.body.style.overflow = 'hidden';
    if (reportDateInput) {
      reportDateInput.focus();
    }
    if (reportDocxStatus) {
      reportDocxStatus.hidden = false;
      reportDocxStatus.textContent = 'Ready.';
    }
    if (reportDocxOutput) {
      reportDocxOutput.value = 'Ready to generate DOCX.';
    }
  }

  function openActivitiesEditorDialog() {
    if (!activitiesEditorDialog) return;
    activitiesEditorDialog.hidden = false;
    if (reportDocxModal) {
      reportDocxModal.classList.add('activities-editor-open');
    }
    if (reportActivitiesEditor) {
      reportActivitiesEditor.focus();
    }
  }

  function closeActivitiesEditorDialog() {
    if (!activitiesEditorDialog) return;
    activitiesEditorDialog.hidden = true;
    if (reportDocxModal) {
      reportDocxModal.classList.remove('activities-editor-open');
    }
  }

  function closeReportDocxModalFn() {
    if (!reportDocxModal) {
      return;
    }
    closeActivitiesEditorDialog();
    reportDocxModal.hidden = true;
    document.body.style.overflow = '';
  }

  async function readTemplateFromFilesFolder() {
    const response = await fetch(encodeURI(REPORT_DOCX_TEMPLATE_PATH));
    if (!response.ok) {
      throw new Error(`Template not found: ${REPORT_DOCX_TEMPLATE_PATH}`);
    }
    return response.arrayBuffer();
  }

  async function downloadReportTemplateAndPromptSelect() {
    const link = document.createElement('a');
    link.href = encodeURI(REPORT_DOCX_TEMPLATE_PATH);
    link.download = 'report maker.docx';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    if (reportDocxStatus) {
      reportDocxStatus.hidden = false;
      reportDocxStatus.textContent = 'Template download started. Please choose the downloaded DOCX file.';
    }
    if (reportDocxTemplateInput) {
      window.setTimeout(() => {
        reportDocxTemplateInput.click();
      }, 250);
    }
  }

  function escapeHtml(value) {
    return String(value || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function insertHtmlAtCaret(html) {
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return false;
    const range = sel.getRangeAt(0);
    range.deleteContents();
    const wrapper = document.createElement('div');
    wrapper.innerHTML = html;
    const frag = document.createDocumentFragment();
    let node = wrapper.firstChild;
    let lastNode = null;
    while (node) {
      lastNode = frag.appendChild(node);
      node = wrapper.firstChild;
    }
    range.insertNode(frag);
    if (lastNode) {
      range.setStartAfter(lastNode);
      range.collapse(true);
      sel.removeAllRanges();
      sel.addRange(range);
    }
    return true;
  }

  function readFileAsDataURL(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(typeof reader.result === 'string' ? reader.result : '');
      reader.onerror = () => reject(new Error(`Failed to read image: ${file.name}`));
      reader.readAsDataURL(file);
    });
  }

  function isActivitiesImageSourceConvertible(src) {
    const value = String(src || '').trim();
    if (!value) return false;
    if (value.startsWith('data:image/')) return false;
    if (value.startsWith('blob:')) return false;
    return /^https?:\/\//i.test(value) || value.startsWith('/');
  }

  async function fetchImageAsDataUrl(src) {
    const attempts = [
      { cache: 'no-store', credentials: 'include', mode: 'cors' },
      { cache: 'no-store', credentials: 'include' },
      { cache: 'no-store' },
    ];

    let lastError = null;
    for (let i = 0; i < attempts.length; i += 1) {
      try {
        const response = await fetch(src, attempts[i]);
        if (!response.ok) {
          throw new Error(`Image fetch failed: ${response.status}`);
        }
        const blob = await response.blob();
        if (!blob || !String(blob.type || '').toLowerCase().startsWith('image/')) {
          throw new Error('Fetched content is not an image.');
        }
        return await new Promise((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = () => resolve(typeof reader.result === 'string' ? reader.result : '');
          reader.onerror = () => reject(new Error('Image conversion failed.'));
          reader.readAsDataURL(blob);
        });
      } catch (error) {
        lastError = error;
      }
    }
    throw lastError || new Error('Image fetch failed.');
  }

  function getFirstSrcFromSrcSet(srcsetValue) {
    const raw = String(srcsetValue || '').trim();
    if (!raw) return '';
    const first = raw.split(',')[0] || '';
    return String(first).trim().split(/\s+/)[0] || '';
  }

  function resolveActivitiesImageSource(imgElement) {
    if (!imgElement) return '';
    const candidates = [
      imgElement.getAttribute('src'),
      imgElement.getAttribute('data-src'),
      imgElement.getAttribute('data-original'),
      getFirstSrcFromSrcSet(imgElement.getAttribute('srcset')),
    ]
      .map((v) => String(v || '').trim())
      .filter(Boolean);
    return candidates[0] || '';
  }

  function sanitizeActivitiesHtml(htmlText) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(String(htmlText || ''), 'text/html');
    const body = doc.body;
    if (!body) return '';

    body.querySelectorAll('script,style,iframe,object,embed').forEach((node) => node.remove());

    body.querySelectorAll('a').forEach((anchor) => {
      const href = String(anchor.getAttribute('href') || '').trim();
      if (href) {
        anchor.setAttribute('href', href);
      }
      anchor.setAttribute('target', '_blank');
      anchor.setAttribute('rel', 'noopener noreferrer');
    });

    body.querySelectorAll('img').forEach((img, idx) => {
      const src = resolveActivitiesImageSource(img);
      if (!src) {
        img.remove();
        return;
      }
      img.setAttribute('src', src);
      img.classList.add('activities-inline-image');
      img.setAttribute('alt', img.getAttribute('alt') || `pasted-image-${Date.now()}-${idx + 1}`);
      img.setAttribute('style', 'max-width:100%;height:auto;border-radius:8px;display:block;margin:6px 0;');
    });

    return body.innerHTML || '';
  }

  async function injectClipboardImagesIntoHtml(safeHtml, imageItems) {
    const html = String(safeHtml || '').trim();
    const items = Array.isArray(imageItems) ? imageItems : [];
    if (!html || items.length === 0) {
      return html;
    }

    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    const imgs = Array.from((doc.body && doc.body.querySelectorAll('img')) || []);
    if (imgs.length === 0) {
      return html;
    }

    let itemIndex = 0;
    for (let i = 0; i < imgs.length; i += 1) {
      if (itemIndex >= items.length) {
        break;
      }
      const file = items[itemIndex] && typeof items[itemIndex].getAsFile === 'function'
        ? items[itemIndex].getAsFile()
        : null;
      if (!file) {
        itemIndex += 1;
        continue;
      }
      try {
        const dataUrl = await readFileAsDataURL(file);
        if (String(dataUrl || '').startsWith('data:image/')) {
          imgs[i].setAttribute('src', dataUrl);
          itemIndex += 1;
        }
      } catch (error) {
        itemIndex += 1;
      }
    }

    return doc.body ? String(doc.body.innerHTML || '') : html;
  }

  async function convertActivitiesEditorExternalImages(options = {}) {
    const replaceOnFail = options.replaceOnFail !== false;
    if (!reportActivitiesEditor) return;
    const imgs = Array.from(reportActivitiesEditor.querySelectorAll('img'));
    for (let i = 0; i < imgs.length; i += 1) {
      const img = imgs[i];
      const src = resolveActivitiesImageSource(img);
      if (!isActivitiesImageSourceConvertible(src)) {
        continue;
      }
      try {
        const dataUrl = await fetchImageAsDataUrl(src);
        if (dataUrl.startsWith('data:image/')) {
          img.setAttribute('src', dataUrl);
        }
      } catch (error) {
        if (replaceOnFail) {
          // Replace external image with readable URL placeholder only during final conversion.
          const placeholder = document.createElement('div');
          placeholder.className = 'activities-image-placeholder';
          placeholder.textContent = `[Image link: ${src}]`;
          const parent = img.parentNode;
          if (parent) {
            parent.replaceChild(placeholder, img);
          }
        } else {
          img.classList.add('activities-inline-image--external');
        }
      }
    }
  }

  async function handleActivitiesPaste(event) {
    if (!reportActivitiesEditor) return;
    const clipboard = event.clipboardData;
    if (!clipboard) return;
    const items = clipboard.items ? Array.from(clipboard.items) : [];
    const imageItems = items.filter((it) => it.type && it.type.startsWith('image/'));
    const htmlData = clipboard.getData ? clipboard.getData('text/html') : '';
    const textData = clipboard.getData ? clipboard.getData('text/plain') : '';
    const hasRichHtml = String(htmlData || '').trim().length > 0;
    const hasImageItems = imageItems.length > 0;

    if (!hasRichHtml && !hasImageItems) {
      return;
    }

    event.preventDefault();

    if (hasRichHtml) {
      let safeHtml = sanitizeActivitiesHtml(htmlData);
      if (safeHtml && hasImageItems) {
        safeHtml = await injectClipboardImagesIntoHtml(safeHtml, imageItems);
      }
      if (safeHtml && insertHtmlAtCaret(safeHtml)) {
        await convertActivitiesEditorExternalImages({ replaceOnFail: false });
        return;
      }
    }

    if (hasImageItems) {
      for (let i = 0; i < imageItems.length; i += 1) {
        const file = imageItems[i].getAsFile();
        if (!file) continue;
        const dataUrl = await readFileAsDataURL(file);
        const altText = `pasted-image-${Date.now()}-${i + 1}`;
        const html = `<div><img class="activities-inline-image" src="${dataUrl}" alt="${escapeHtml(altText)}" style="max-width:100%;height:auto;border-radius:8px;display:block;margin:6px 0;" /></div>`;
        if (!insertHtmlAtCaret(html)) {
          reportActivitiesEditor.insertAdjacentHTML('beforeend', html);
        }
      }
      return;
    }

    if (String(textData || '').trim()) {
      insertHtmlAtCaret(escapeHtml(textData).replace(/\n/g, '<br/>'));
    }
  }

  function getActivitiesEditorText() {
    if (!reportActivitiesEditor) return '';
    const clone = reportActivitiesEditor.cloneNode(true);
    const imgs = clone.querySelectorAll('img');
    imgs.forEach((img, idx) => {
      const marker = document.createTextNode(`[Image ${idx + 1}]`);
      img.replaceWith(marker);
    });
    const text = String(clone.innerText || '')
      .replace(/\u00a0/g, ' ')
      .replace(/\r/g, '')
      .replace(/\n{3,}/g, '\n\n')
      .trim();
    return text;
  }

  function getActivitiesEditorImages() {
    if (!reportActivitiesEditor) return [];
    const imgs = Array.from(reportActivitiesEditor.querySelectorAll('img'));
    return imgs
      .map((img, index) => {
        const src = String(img.getAttribute('src') || '');
        if (!src.startsWith('data:image/')) {
          return null;
        }
        return {
          src,
          name: `activities_image_${index + 1}`,
        };
      })
      .filter(Boolean);
  }

  function applyActivitiesEditorCommand(commandName) {
    if (!reportActivitiesEditor) return;
    reportActivitiesEditor.focus();
    try {
      document.execCommand(commandName, false, null);
    } catch (error) {
      // Keep editor usable even if browser blocks formatting command.
    }
  }

  function applyActivitiesBulletsByLine() {
    if (!reportActivitiesEditor) return;
    reportActivitiesEditor.focus();
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) {
      applyActivitiesEditorCommand('insertUnorderedList');
      return;
    }

    const range = selection.getRangeAt(0);
    const editorRange = document.createRange();
    editorRange.selectNodeContents(reportActivitiesEditor);
    const startsInside = range.compareBoundaryPoints(Range.START_TO_START, editorRange) >= 0;
    const endsInside = range.compareBoundaryPoints(Range.END_TO_END, editorRange) <= 0;
    if (!startsInside || !endsInside) {
      applyActivitiesEditorCommand('insertUnorderedList');
      return;
    }

    const selectedText = String(selection.toString() || '');
    if (!selectedText.trim()) {
      applyActivitiesEditorCommand('insertUnorderedList');
      return;
    }

    const lines = selectedText
      .replace(/\r/g, '')
      .split('\n')
      .map((line) => line.trim())
      .filter((line) => line.length > 0);

    if (lines.length <= 1) {
      applyActivitiesEditorCommand('insertUnorderedList');
      return;
    }

    const ul = document.createElement('ul');
    lines.forEach((line) => {
      const li = document.createElement('li');
      li.textContent = line;
      ul.appendChild(li);
    });

    range.deleteContents();
    range.insertNode(ul);

    const after = document.createRange();
    after.setStartAfter(ul);
    after.collapse(true);
    selection.removeAllRanges();
    selection.addRange(after);
  }


  function getActivitiesBlocksFromEditor() {
    if (!reportActivitiesEditor) return [];

    const blocks = [];
    const pushText = (rawText, asBullet) => {
      const lines = String(rawText || '')
        .replace(/\u00a0/g, ' ')
        .replace(/\r/g, '')
        .split('\n')
        .map((line) => line.trim())
        .filter((line) => line.length > 0);
      lines.forEach((line) => {
        blocks.push(asBullet ? { type: 'bullet', text: line } : { type: 'text', text: line });
      });
    };

    const walk = (node, inBullet) => {
      if (!node) return;
      if (node.nodeType === Node.TEXT_NODE) {
        pushText(node.textContent || '', inBullet);
        return;
      }
      if (node.nodeType !== Node.ELEMENT_NODE) {
        return;
      }

      const el = node;
      const tag = String(el.tagName || '').toLowerCase();

      if (tag === 'img') {
        const src = String(el.getAttribute('src') || '');
        if (src.startsWith('data:image/')) {
          blocks.push({
            type: 'image',
            image: {
              src,
              name: `activities_image_${blocks.length + 1}`,
            },
          });
        } else if (src) {
          blocks.push({ type: 'text', text: `[Image link: ${src}]` });
        }
        return;
      }

      if (tag === 'a') {
        const href = String(el.getAttribute('href') || '').trim();
        const text = normalizeText(el.textContent || '');
        const withHref = href && text && !text.includes(href)
          ? `${text} (${href})`
          : (text || href);
        if (withHref) {
          blocks.push(inBullet ? { type: 'bullet', text: withHref } : { type: 'text', text: withHref });
        }
        return;
      }

      if (tag === 'ul' || tag === 'ol') {
        Array.from(el.children).forEach((child) => {
          if (String(child.tagName || '').toLowerCase() === 'li') {
            walk(child, true);
          } else {
            walk(child, inBullet);
          }
        });
        return;
      }

      if (tag === 'li') {
        if (el.childNodes.length === 0) {
          return;
        }
        Array.from(el.childNodes).forEach((child) => walk(child, true));
        return;
      }

      if (tag === 'br') {
        blocks.push(inBullet ? { type: 'bullet', text: '' } : { type: 'text', text: '' });
        return;
      }

      Array.from(el.childNodes).forEach((child) => walk(child, inBullet));
    };

    Array.from(reportActivitiesEditor.childNodes).forEach((node) => walk(node, false));

    return blocks.filter((block) => {
      if (!block) return false;
      if (block.type === 'image') return Boolean(block.image && block.image.src);
      return String(block.text || '').trim().length > 0;
    });
  }

  function xmlEscape(value) {
    return String(value || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function makeWordRun(text, bold) {
    const safeText = xmlEscape(text);
    const boldTag = bold ? '<w:b/>' : '';
    return `<w:r><w:rPr>${boldTag}<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">${safeText}</w:t></w:r>`;
  }

  function makeWordParagraph(text, bold) {
    return `<w:p><w:pPr><w:spacing w:line="300" w:lineRule="auto"/></w:pPr>${makeWordRun(text, bold)}</w:p>`;
  }

  function makeInfoRow(label, value, labelBold, valueBold) {
    return [
      '<w:tr>',
      '<w:tc><w:tcPr><w:tcW w:w="2800" w:type="dxa"/></w:tcPr>',
      makeWordParagraph(label, labelBold),
      '</w:tc>',
      '<w:tc><w:tcPr><w:tcW w:w="7200" w:type="dxa"/></w:tcPr>',
      makeWordParagraph(value, valueBold),
      '</w:tc>',
      '</w:tr>',
    ].join('');
  }

  function buildReportDocxXml(payload) {
    const contentLines = String(payload.body || '').split(/\r?\n/);
    const contentParagraphs = contentLines.length > 0
      ? contentLines.map((line) => makeWordParagraph(line || ' ', false)).join('')
      : makeWordParagraph(' ', false);

    const tableXml = [
      '<w:tbl>',
      '<w:tblPr>',
      '<w:tblW w:w="10000" w:type="dxa"/>',
      '<w:tblBorders>',
      '<w:top w:val="single" w:sz="8" w:color="DBDBDB"/>',
      '<w:left w:val="single" w:sz="8" w:color="DBDBDB"/>',
      '<w:bottom w:val="single" w:sz="8" w:color="DBDBDB"/>',
      '<w:right w:val="single" w:sz="8" w:color="DBDBDB"/>',
      '<w:insideH w:val="single" w:sz="8" w:color="DBDBDB"/>',
      '<w:insideV w:val="single" w:sz="8" w:color="DBDBDB"/>',
      '</w:tblBorders>',
      '</w:tblPr>',
      makeInfoRow('DATE:', payload.date, true, true),
      makeInfoRow('JOF #', payload.jof, true, false),
      makeInfoRow('Invoice No.', payload.invoice, true, false),
      makeInfoRow('Client Name:', payload.client, true, false),
      makeInfoRow('Report Title:', payload.title, true, false),
      '</w:tbl>',
    ].join('');

    return [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
      '<w:body>',
      tableXml,
      makeWordParagraph(' ', false),
      makeWordParagraph('REPORT DETAILS', true),
      contentParagraphs,
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>',
      '</w:body>',
      '</w:document>',
    ].join('');
  }

  async function buildReportDocxBlob(payload) {
    if (typeof JSZip === 'undefined') {
      throw new Error('JSZip not loaded');
    }
    const zip = new JSZip();
    zip.file('[Content_Types].xml', [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
      '<Default Extension="xml" ContentType="application/xml"/>',
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
      '</Types>',
    ].join(''));

    zip.folder('_rels').file('.rels', [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>',
      '</Relationships>',
    ].join(''));

    zip.folder('word').file('document.xml', buildReportDocxXml(payload));
    zip.folder('word').folder('_rels').file('document.xml.rels', [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>',
    ].join(''));

    return zip.generateAsync({ type: 'blob' });
  }

  function replaceFirstTextMatchInTemplate(docXml, matcher, replacement) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(docXml, 'application/xml');
    const nodes = Array.from(xmlDoc.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 't'));
    let replaced = false;
    nodes.forEach((node) => {
      if (replaced) return;
      const text = String(node.textContent || '');
      if (matcher(text)) {
        node.textContent = replacement;
        replaced = true;
      }
    });
    const serializer = new XMLSerializer();
    return serializer.serializeToString(xmlDoc);
  }

  function getWNodeList(parent, tagName) {
    return Array.from(parent.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', tagName));
  }

  function getClosestByLocalName(node, localName) {
    let current = node;
    while (current) {
      if (current.localName === localName) {
        return current;
      }
      current = current.parentNode;
    }
    return null;
  }

  function getNextSiblingByLocalName(node, localName) {
    if (!node || !node.parentNode) return null;
    let sibling = node.nextSibling;
    while (sibling) {
      if (sibling.nodeType === 1 && sibling.localName === localName) {
        return sibling;
      }
      sibling = sibling.nextSibling;
    }
    return null;
  }

  const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
  const PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture';

  function createMinimalTextParagraph(xmlDoc, textValue, options = {}) {
    const isBold = Boolean(options.bold);
    const p = xmlDoc.createElementNS(W_NS, 'w:p');
    const r = xmlDoc.createElementNS(W_NS, 'w:r');
    const rPr = xmlDoc.createElementNS(W_NS, 'w:rPr');
    const rFonts = xmlDoc.createElementNS(W_NS, 'w:rFonts');
    rFonts.setAttribute('w:ascii', 'Arial');
    rFonts.setAttribute('w:hAnsi', 'Arial');
    const sz = xmlDoc.createElementNS(W_NS, 'w:sz');
    sz.setAttribute('w:val', '24');
    const szCs = xmlDoc.createElementNS(W_NS, 'w:szCs');
    szCs.setAttribute('w:val', '24');
    rPr.appendChild(rFonts);
    if (isBold) {
      const b = xmlDoc.createElementNS(W_NS, 'w:b');
      rPr.appendChild(b);
    }
    rPr.appendChild(sz);
    rPr.appendChild(szCs);
    const t = xmlDoc.createElementNS(W_NS, 'w:t');
    t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
    t.textContent = String(textValue || ' ');
    r.appendChild(rPr);
    r.appendChild(t);
    p.appendChild(r);
    return p;
  }

  function ensureCellHasTextNode(cellNode, xmlDoc) {
    const textNodes = getWNodeList(cellNode, 't');
    if (textNodes.length > 0) return textNodes;
    cellNode.appendChild(createMinimalTextParagraph(xmlDoc, ' '));
    return getWNodeList(cellNode, 't');
  }

  function clearCellContent(cellNode, xmlDoc) {
    if (!cellNode) return;
    while (cellNode.firstChild) {
      cellNode.removeChild(cellNode.firstChild);
    }
    cellNode.appendChild(createMinimalTextParagraph(xmlDoc, ' '));
  }

  function setCellText(cellNode, value, xmlDoc, options = {}) {
    if (!cellNode) return false;
    const isBold = Boolean(options.bold);
    while (cellNode.firstChild) {
      cellNode.removeChild(cellNode.firstChild);
    }
    cellNode.appendChild(createMinimalTextParagraph(xmlDoc, String(value || ' '), { bold: isBold }));
    return true;
  }

  function normalizeText(value) {
    return String(value || '').replace(/\s+/g, ' ').trim();
  }

  function formatReportDateMmDdYyyy(rawValue) {
    const raw = String(rawValue || '').trim();
    if (!raw) return '';
    const ymd = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (ymd) {
      return `${ymd[2]}/${ymd[3]}/${ymd[1]}`;
    }
    const mdy = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (mdy) {
      const mm = mdy[1].padStart(2, '0');
      const dd = mdy[2].padStart(2, '0');
      return `${mm}/${dd}/${mdy[3]}`;
    }
    return raw;
  }

  function isTruthyText(value) {
    return normalizeText(value).length > 0;
  }

  function getTextFromNode(node) {
    const textNodes = getWNodeList(node, 't');
    return normalizeText(textNodes.map((t) => String(t.textContent || '')).join(' '));
  }

  function setValueByLabelInNextCell(xmlDoc, labelText, value, options = {}) {
    if (!isTruthyText(value)) return false;
    const textNodes = getWNodeList(xmlDoc, 't');
    for (let i = 0; i < textNodes.length; i += 1) {
      const text = String(textNodes[i].textContent || '').trim();
      if (!text.includes(labelText)) continue;
      const tc = getClosestByLocalName(textNodes[i], 'tc');
      const nextTc = getNextSiblingByLocalName(tc, 'tc');
      if (nextTc && setCellText(nextTc, value, xmlDoc, options)) {
        return true;
      }
    }
    return false;
  }

  function setValueInRowBelowLabel(xmlDoc, labelText, value, options = {}) {
    if (!isTruthyText(value)) return false;
    const textNodes = getWNodeList(xmlDoc, 't');
    for (let i = 0; i < textNodes.length; i += 1) {
      const text = String(textNodes[i].textContent || '').trim();
      if (!text.includes(labelText)) continue;
      const tr = getClosestByLocalName(textNodes[i], 'tr');
      const nextTr = getNextSiblingByLocalName(tr, 'tr');
      if (!nextTr) continue;
      const firstTc = Array.from(nextTr.childNodes).find((n) => n.nodeType === 1 && n.localName === 'tc');
      if (firstTc && setCellText(firstTc, value, xmlDoc, options)) {
        return true;
      }
    }
    return false;
  }

  function setValueAfterLabelInSameCell(xmlDoc, labelText, value, options = {}) {
    if (!isTruthyText(value)) return false;
    const textNodes = getWNodeList(xmlDoc, 't');
    for (let i = 0; i < textNodes.length; i += 1) {
      const text = String(textNodes[i].textContent || '');
      if (!text.includes(labelText)) continue;
      const tc = getClosestByLocalName(textNodes[i], 'tc');
      if (tc) {
        return setCellText(tc, `${labelText} ${String(value).trim()}`, xmlDoc, options);
      }
    }
    return false;
  }

  function getRowBelowLabelCell(xmlDoc, labelText) {
    const textNodes = getWNodeList(xmlDoc, 't');
    for (let i = 0; i < textNodes.length; i += 1) {
      const text = String(textNodes[i].textContent || '').trim();
      if (!text.includes(labelText)) continue;
      const tr = getClosestByLocalName(textNodes[i], 'tr');
      const nextTr = getNextSiblingByLocalName(tr, 'tr');
      if (!nextTr) continue;
      const firstTc = Array.from(nextTr.childNodes).find((n) => n.nodeType === 1 && n.localName === 'tc');
      if (firstTc) return firstTc;
    }
    return null;
  }

  function appendTextLinesToCell(cellNode, xmlDoc, textValue) {
    if (!cellNode) return;
    clearCellContent(cellNode, xmlDoc);
    while (cellNode.firstChild) {
      cellNode.removeChild(cellNode.firstChild);
    }
    const lines = String(textValue || '')
      .replace(/\r/g, '')
      .split('\n');
    if (lines.length === 0) {
      cellNode.appendChild(createMinimalTextParagraph(xmlDoc, ' '));
      return;
    }
    lines.forEach((line) => {
      const safeLine = String(line || '').replace(/\[\[\[BULLET\]\]\]/g, '').trim();
      const normalized = safeLine;
      const isTsActivityHeading = /^TS\s+ACTIVITY\b/i.test(normalized);
      cellNode.appendChild(createMinimalTextParagraph(xmlDoc, safeLine || ' ', { bold: isTsActivityHeading }));
      if (isTsActivityHeading) {
        // Keep two newlines after TS ACTIVITY heading.
        cellNode.appendChild(createMinimalTextParagraph(xmlDoc, ' ', { bold: false }));
      }
    });
  }

  function appendActivitiesBlocksToCell(cellNode, xmlDoc, blocks, renderImageBlock) {
    if (!cellNode) return;
    clearCellContent(cellNode, xmlDoc);
    while (cellNode.firstChild) {
      cellNode.removeChild(cellNode.firstChild);
    }

    let hasContent = false;
    (Array.isArray(blocks) ? blocks : []).forEach((block, i) => {
      if (!block) return;
      if (block.type === 'text') {
        const lines = String(block.text || '').replace(/\r/g, '').split('\n');
        lines.forEach((line) => {
          const safeLine = String(line || '').replace(/\[\[\[BULLET\]\]\]/g, '').trim();
          const normalized = safeLine;
          const isTsActivityHeading = /^TS\s+ACTIVITY\b/i.test(normalized);
          cellNode.appendChild(createMinimalTextParagraph(xmlDoc, safeLine || ' ', { bold: isTsActivityHeading }));
          if (isTsActivityHeading) {
            // Keep two newlines after TS ACTIVITY heading.
            cellNode.appendChild(createMinimalTextParagraph(xmlDoc, ' ', { bold: false }));
          }
          hasContent = true;
        });
        return;
      }
      if (block.type === 'bullet') {
        const bulletText = String(block.text || '').trim();
        cellNode.appendChild(createMinimalTextParagraph(xmlDoc, `• ${bulletText || ' '}`, { bold: false }));
        hasContent = true;
        return;
      }
      if (block.type === 'image' && typeof renderImageBlock === 'function') {
        const node = renderImageBlock(block.image, i);
        if (node) {
          cellNode.appendChild(node);
          hasContent = true;
        }
      }
    });

    if (!hasContent) {
      cellNode.appendChild(createMinimalTextParagraph(xmlDoc, ' '));
    }
  }

  function parseDataUrlImage(dataUrl) {
    const match = String(dataUrl || '').match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
    if (!match) return null;
    const mime = match[1].toLowerCase();
    const base64 = match[2];
    let ext = 'png';
    if (mime.includes('jpeg') || mime.includes('jpg')) ext = 'jpg';
    if (mime.includes('png')) ext = 'png';
    return { mime, base64, ext };
  }

  async function readDataUrlImageSize(dataUrl) {
    const src = String(dataUrl || '').trim();
    if (!src.startsWith('data:image/')) {
      return null;
    }
    return await new Promise((resolve) => {
      const img = new Image();
      img.onload = () => {
        const width = Number(img.naturalWidth || img.width || 0);
        const height = Number(img.naturalHeight || img.height || 0);
        if (width > 0 && height > 0) {
          resolve({ width, height });
        } else {
          resolve(null);
        }
      };
      img.onerror = () => resolve(null);
      img.src = src;
    });
  }

  function addDocxImageRelationship(relsDoc, relId, target) {
    const relsRoot = relsDoc.documentElement;
    const relNode = relsDoc.createElementNS(REL_NS, 'Relationship');
    relNode.setAttribute('Id', relId);
    relNode.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
    relNode.setAttribute('Target', target);
    relsRoot.appendChild(relNode);
  }

  function buildImageParagraphNode(xmlDoc, relId, docPrId, name, cx, cy) {
    const p = xmlDoc.createElementNS(W_NS, 'w:p');
    const r = xmlDoc.createElementNS(W_NS, 'w:r');
    const drawing = xmlDoc.createElementNS(W_NS, 'w:drawing');
    const inline = xmlDoc.createElementNS(WP_NS, 'wp:inline');
    inline.setAttribute('distT', '0');
    inline.setAttribute('distB', '0');
    inline.setAttribute('distL', '0');
    inline.setAttribute('distR', '0');

    const extent = xmlDoc.createElementNS(WP_NS, 'wp:extent');
    extent.setAttribute('cx', String(cx));
    extent.setAttribute('cy', String(cy));
    inline.appendChild(extent);

    const effectExtent = xmlDoc.createElementNS(WP_NS, 'wp:effectExtent');
    effectExtent.setAttribute('l', '0');
    effectExtent.setAttribute('t', '0');
    effectExtent.setAttribute('r', '0');
    effectExtent.setAttribute('b', '0');
    inline.appendChild(effectExtent);

    const docPr = xmlDoc.createElementNS(WP_NS, 'wp:docPr');
    docPr.setAttribute('id', String(docPrId));
    docPr.setAttribute('name', name);
    inline.appendChild(docPr);

    const cNvGraphicFramePr = xmlDoc.createElementNS(WP_NS, 'wp:cNvGraphicFramePr');
    const graphicFrameLocks = xmlDoc.createElementNS(A_NS, 'a:graphicFrameLocks');
    graphicFrameLocks.setAttribute('noChangeAspect', '1');
    cNvGraphicFramePr.appendChild(graphicFrameLocks);
    inline.appendChild(cNvGraphicFramePr);

    const graphic = xmlDoc.createElementNS(A_NS, 'a:graphic');
    const graphicData = xmlDoc.createElementNS(A_NS, 'a:graphicData');
    graphicData.setAttribute('uri', 'http://schemas.openxmlformats.org/drawingml/2006/picture');

    const pic = xmlDoc.createElementNS(PIC_NS, 'pic:pic');
    const nvPicPr = xmlDoc.createElementNS(PIC_NS, 'pic:nvPicPr');
    const cNvPr = xmlDoc.createElementNS(PIC_NS, 'pic:cNvPr');
    cNvPr.setAttribute('id', '0');
    cNvPr.setAttribute('name', name);
    nvPicPr.appendChild(cNvPr);
    nvPicPr.appendChild(xmlDoc.createElementNS(PIC_NS, 'pic:cNvPicPr'));
    pic.appendChild(nvPicPr);

    const blipFill = xmlDoc.createElementNS(PIC_NS, 'pic:blipFill');
    const blip = xmlDoc.createElementNS(A_NS, 'a:blip');
    blip.setAttributeNS(R_NS, 'r:embed', relId);
    blipFill.appendChild(blip);
    const stretch = xmlDoc.createElementNS(A_NS, 'a:stretch');
    stretch.appendChild(xmlDoc.createElementNS(A_NS, 'a:fillRect'));
    blipFill.appendChild(stretch);
    pic.appendChild(blipFill);

    const spPr = xmlDoc.createElementNS(PIC_NS, 'pic:spPr');
    const xfrm = xmlDoc.createElementNS(A_NS, 'a:xfrm');
    const off = xmlDoc.createElementNS(A_NS, 'a:off');
    off.setAttribute('x', '0');
    off.setAttribute('y', '0');
    const ext = xmlDoc.createElementNS(A_NS, 'a:ext');
    ext.setAttribute('cx', String(cx));
    ext.setAttribute('cy', String(cy));
    xfrm.appendChild(off);
    xfrm.appendChild(ext);
    spPr.appendChild(xfrm);
    const prstGeom = xmlDoc.createElementNS(A_NS, 'a:prstGeom');
    prstGeom.setAttribute('prst', 'rect');
    prstGeom.appendChild(xmlDoc.createElementNS(A_NS, 'a:avLst'));
    spPr.appendChild(prstGeom);
    pic.appendChild(spPr);

    graphicData.appendChild(pic);
    graphic.appendChild(graphicData);
    inline.appendChild(graphic);
    drawing.appendChild(inline);
    r.appendChild(drawing);
    p.appendChild(r);
    return p;
  }

  function splitReportBodySections(bodyText) {
    const raw = String(bodyText || '').replace(/\r/g, '');
    if (!raw.trim()) {
      return {
        concern: '',
        activities: '',
        rootCause: '',
        preventiveAction: '',
        nextSteps: '',
        currentStatus: '',
      };
    }

    const lines = raw.split('\n');
    const result = {
      concern: '',
      activities: raw.trim(),
      rootCause: '',
      preventiveAction: '',
      nextSteps: '',
      currentStatus: '',
    };

    const markers = [
      { key: 'concern', labels: ['concern/s:', 'concern:'] },
      { key: 'activities', labels: ['activities of ts:', 'activities:'] },
      { key: 'rootCause', labels: ['root cause:'] },
      { key: 'preventiveAction', labels: ['preventive action:'] },
      { key: 'nextSteps', labels: ['next steps:'] },
      { key: 'currentStatus', labels: ['current status:'] },
    ];

    let currentKey = '';
    const bucket = {
      concern: [],
      activities: [],
      rootCause: [],
      preventiveAction: [],
      nextSteps: [],
      currentStatus: [],
    };

    lines.forEach((line) => {
      const normalized = normalizeText(line).toLowerCase();
      const found = markers.find((m) => m.labels.some((label) => normalized.startsWith(label)));
      if (found) {
        currentKey = found.key;
        const cutLabel = found.labels.find((label) => normalized.startsWith(label)) || '';
        const inlineValue = line.slice(line.toLowerCase().indexOf(cutLabel) + cutLabel.length).trim();
        if (inlineValue) {
          bucket[currentKey].push(inlineValue);
        }
        return;
      }
      if (currentKey) {
        bucket[currentKey].push(line);
      }
    });

    Object.keys(bucket).forEach((key) => {
      const joined = bucket[key].join('\n').trim();
      if (joined) {
        result[key] = joined;
      }
    });

    if (!result.activities && raw.trim()) {
      result.activities = raw.trim();
    }

    return result;
  }

  function applyTemplateSectionValues(xmlDoc, payload) {
    const bodySections = splitReportBodySections(payload.body || '');
    const concernValue = isTruthyText(payload.concern) ? payload.concern : (isTruthyText(payload.title) ? payload.title : bodySections.concern);
    const rootCauseValue = isTruthyText(payload.rootCause) ? payload.rootCause : bodySections.rootCause;
    const preventiveValue = isTruthyText(payload.preventiveAction) ? payload.preventiveAction : bodySections.preventiveAction;
    const nextStepsValue = isTruthyText(payload.nextSteps) ? payload.nextSteps : bodySections.nextSteps;
    const currentStatusValue = isTruthyText(payload.currentStatus) ? payload.currentStatus : bodySections.currentStatus;

    setValueInRowBelowLabel(xmlDoc, 'Concern/s:', concernValue, { bold: true });
    setValueInRowBelowLabel(xmlDoc, 'Root Cause:', rootCauseValue, { bold: true });
    setValueInRowBelowLabel(xmlDoc, 'Preventive Action:', preventiveValue, { bold: true });

    // In this template, Next Steps is usually in the same cell as the label.
    if (!setValueAfterLabelInSameCell(xmlDoc, 'Next Steps:', nextStepsValue, { bold: true })) {
      setValueInRowBelowLabel(xmlDoc, 'Next Steps:', nextStepsValue, { bold: true });
    }

    setValueInRowBelowLabel(xmlDoc, 'Current Status:', currentStatusValue, { bold: true });
  }

  function buildDocxFromTemplateXml(templateXml, payload) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(templateXml, 'application/xml');

    setValueByLabelInNextCell(xmlDoc, 'DATE:', payload.date, { bold: true });
    setValueAfterLabelInSameCell(xmlDoc, 'JOF #', payload.jof, { bold: true });
    setValueAfterLabelInSameCell(xmlDoc, 'Invoice No.', payload.invoice, { bold: true });
    setValueByLabelInNextCell(xmlDoc, 'Client Name:', payload.client, { bold: true });
    setValueByLabelInNextCell(xmlDoc, 'TS Assigned:', payload.tsAssigned, { bold: true });

    applyTemplateSectionValues(xmlDoc, payload);

    return xmlDoc;
  }

  async function buildReportDocxBlobFromTemplate(templateBuffer, payload) {
    if (typeof JSZip === 'undefined') {
      throw new Error('JSZip not loaded');
    }
    const zip = await JSZip.loadAsync(templateBuffer);
    const docFile = zip.file('word/document.xml');
    const relsFile = zip.file('word/_rels/document.xml.rels');
    if (!docFile) {
      throw new Error('Template document.xml not found.');
    }
    if (!relsFile) {
      throw new Error('Template document.xml.rels not found.');
    }
    const docXml = await docFile.async('string');
    const relsXml = await relsFile.async('string');
    const parser = new DOMParser();
    const relsDoc = parser.parseFromString(relsXml, 'application/xml');
    const xmlDoc = buildDocxFromTemplateXml(docXml, payload);

    const imageSizeMap = new Map();
    const imageBlocks = Array.isArray(payload.activitiesBlocks) ? payload.activitiesBlocks : [];
    const fallbackImages = Array.isArray(payload.activitiesImages) ? payload.activitiesImages : [];
    const allImageEntries = [];
    imageBlocks.forEach((block) => {
      if (block && block.type === 'image' && block.image && block.image.src) {
        allImageEntries.push(block.image);
      }
    });
    fallbackImages.forEach((img) => {
      if (img && img.src) {
        allImageEntries.push(img);
      }
    });
    const uniqueImageSrcs = Array.from(new Set(allImageEntries.map((entry) => String(entry.src || '')).filter(Boolean)));
    for (let i = 0; i < uniqueImageSrcs.length; i += 1) {
      const src = uniqueImageSrcs[i];
      const size = await readDataUrlImageSize(src);
      if (size) {
        imageSizeMap.set(src, size);
      }
    }

    const activityCell = getRowBelowLabelCell(xmlDoc, 'Activities of TS:');
    let relIdCounter = 1000;
    let docPrCounter = 1000;
    const renderImageBlock = (img, index) => {
      const parsed = parseDataUrlImage(img.src);
      if (!parsed) return null;
      const relId = `rIdActivitiesImg${relIdCounter + index}`;
      const fileName = `activities_image_${Date.now()}_${index + 1}.${parsed.ext}`;
      const target = `media/${fileName}`;
      addDocxImageRelationship(relsDoc, relId, target);
      zip.file(`word/${target}`, parsed.base64, { base64: true });
      const EMU_PER_PX = 9525;
      const MAX_CX = 4876800; // ~5.12in, fits inside template cell
      let cx = MAX_CX;
      let cy = 2743200;
      const srcKey = String(img && img.src ? img.src : '');
      const size = srcKey ? imageSizeMap.get(srcKey) : null;
      if (size && size.width > 0 && size.height > 0) {
        cx = Math.round(size.width * EMU_PER_PX);
        cy = Math.round(size.height * EMU_PER_PX);
        if (cx > MAX_CX) {
          const ratio = MAX_CX / cx;
          cx = MAX_CX;
          cy = Math.max(1, Math.round(cy * ratio));
        }
      }
      return buildImageParagraphNode(
        xmlDoc,
        relId,
        docPrCounter + index,
        img.name || `Activities Image ${index + 1}`,
        cx,
        cy,
      );
    };

    if (activityCell) {
      const blocks = Array.isArray(payload.activitiesBlocks) ? payload.activitiesBlocks : [];
      if (blocks.length > 0) {
        appendActivitiesBlocksToCell(activityCell, xmlDoc, blocks, renderImageBlock);
      } else if (isTruthyText(payload.activities)) {
        appendTextLinesToCell(activityCell, xmlDoc, payload.activities);
        const images = Array.isArray(payload.activitiesImages) ? payload.activitiesImages : [];
        images.forEach((img, index) => {
          const node = renderImageBlock(img, index);
          if (node) {
            activityCell.appendChild(node);
          }
        });
      }
    }

    const serializer = new XMLSerializer();
    zip.file('word/document.xml', serializer.serializeToString(xmlDoc));
    zip.file('word/_rels/document.xml.rels', serializer.serializeToString(relsDoc));
    return zip.generateAsync({ type: 'blob' });
  }

  function downloadDocxBlob(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.setTimeout(() => URL.revokeObjectURL(url), 1200);
  }

  function closeXmlRenameModalFn() {
    if (!xmlRenameModal) {
      return;
    }
    xmlRenameModal.hidden = true;
    document.body.style.overflow = '';
  }

  function readTextFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(typeof reader.result === 'string' ? reader.result : '');
      reader.onerror = () => reject(new Error(`Failed to read file: ${file.name}`));
      reader.readAsText(file);
    });
  }

  function updateXmlTagValue(xmlText, tagName, newValue) {
    const escapedTag = String(tagName || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const openCloseRegex = new RegExp(`<${escapedTag}>[\\s\\S]*?<\\/${escapedTag}>`, 'gi');
    const replacement = `<${tagName}>${newValue}</${tagName}>`;
    if (openCloseRegex.test(xmlText)) {
      return xmlText.replace(openCloseRegex, replacement);
    }
    return xmlText;
  }

  function buildRenamedXmlFileName(originalName, tenantIdValue) {
    const dotIndex = originalName.lastIndexOf('.');
    const base = dotIndex >= 0 ? originalName.slice(0, dotIndex) : originalName;
    const ext = dotIndex >= 0 ? originalName.slice(dotIndex) : '.xml';
    const salesPattern = /^(sales_)[^_]+(_.*)$/i;
    if (salesPattern.test(base)) {
      return base.replace(salesPattern, `$1${tenantIdValue}$2`) + ext;
    }
    return `${base}_${tenantIdValue}${ext}`;
  }

  function downloadXmlBlob(content, fileName) {
    const blob = new Blob([content], { type: 'application/xml;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function openAllScriptModal() {
    if (!allScriptModal) {
      return;
    }
    renderAllScriptOutput();
    allScriptModal.hidden = false;
    document.body.style.overflow = 'hidden';
  }

  function closeAllScriptModalFn() {
    if (!allScriptModal) {
      return;
    }
    allScriptModal.hidden = true;
    document.body.style.overflow = '';
  }

  function hideFrecnoCopyButton() {
    if (copyFrecnoCodeBtn) {
      copyFrecnoCodeBtn.disabled = true;
    }
    if (copyFrecnoCodeRow) {
      copyFrecnoCodeRow.hidden = true;
    }
  }

  function openAyalaModal() {
    if (!ayalaModal) {
      return;
    }
    ayalaModal.hidden = false;
    document.body.style.overflow = 'hidden';
    setAyalaPanel('checker');
    if (hourlyCsvInput) {
      hourlyCsvInput.focus();
    }
    if (ayalaStatus) {
      ayalaStatus.hidden = false;
      ayalaStatus.textContent = 'Ayala checker ready. Step 1: choose CSV hourly files then click Analyze Hourly.';
    }
    resetAyalaReportOutput();
    setAyalaHourlyFirstMode(ayalaHourlyAnalyzed);
  }

  function setAyalaPanel(viewName) {
    const isChecker = viewName === 'checker';
    if (ayalaCheckerPanel) {
      ayalaCheckerPanel.hidden = !isChecker;
    }
    if (ayalaGuidelinesPanel) {
      ayalaGuidelinesPanel.hidden = isChecker;
    }
    if (ayalaCheckerTabBtn) {
      ayalaCheckerTabBtn.classList.toggle('active', isChecker);
    }
    if (ayalaGuidelinesTabBtn) {
      ayalaGuidelinesTabBtn.classList.toggle('active', !isChecker);
    }
  }

  function triggerAyalaGuidelinesDirectDownload() {
    const downloadLink = document.createElement('a');
    downloadLink.href = encodeURI(ayalaGuidelinesZipPath);
    downloadLink.download = 'Ayala guidelines.zip';
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
  }

  async function downloadAyalaGuidelinesZip() {
    if (isAyalaDownloadInProgress) {
      return;
    }

    isAyalaDownloadInProgress = true;
    if (ayalaGuidelinesTabBtn) {
      ayalaGuidelinesTabBtn.disabled = true;
    }
    if (ayalaStatus) {
      ayalaStatus.hidden = false;
      ayalaStatus.textContent = 'Preparing Ayala guidelines ZIP...';
    }

    try {
      if (ayalaStatus) {
        ayalaStatus.textContent = 'Downloading Ayala guidelines ZIP...';
      }

      // `fetch` gives us stronger success/fail signaling when hosted via http(s).
      const response = await fetch(encodeURI(ayalaGuidelinesZipPath));
      if (!response.ok) {
        throw new Error(`HTTP ${response.status}`);
      }

      const zipBlob = await response.blob();
      const downloadUrl = URL.createObjectURL(zipBlob);
      const downloadLink = document.createElement('a');
      downloadLink.href = downloadUrl;
      downloadLink.download = 'Ayala guidelines.zip';
      document.body.appendChild(downloadLink);
      downloadLink.click();
      document.body.removeChild(downloadLink);
      window.setTimeout(() => URL.revokeObjectURL(downloadUrl), 1000);

      if (ayalaStatus) {
        ayalaStatus.hidden = false;
        ayalaStatus.textContent = 'Download successfully.';
      }
    } catch (error) {
      // Fallback for environments where fetch to local/static files is blocked (e.g. file://).
      triggerAyalaGuidelinesDirectDownload();
      if (ayalaStatus) {
        ayalaStatus.hidden = false;
        ayalaStatus.textContent = 'Download successfully.';
      }
    } finally {
      isAyalaDownloadInProgress = false;
      if (ayalaGuidelinesTabBtn) {
        ayalaGuidelinesTabBtn.disabled = false;
      }
    }
  }

  function closeAyalaModalFn() {
    if (!ayalaModal) {
      return;
    }
    ayalaModal.hidden = true;
    document.body.style.overflow = '';
  }

  function openAyalaResultModal(message, isError) {
    if (!ayalaResultModal || !ayalaResultText) {
      return;
    }
    const rawMessage = String(message || '').trim();
    const classifyLineTone = (line) => {
      const text = String(line || '').trim();
      if (!text) return 'neutral';
      const lower = text.toLowerCase();
      if (
        lower.includes('cross validation, discrepancy')
        || lower.includes('discrepancy (')
        || lower.includes(' mismatch')
        || lower.includes(' diff ')
        || lower.includes('duplicate transaction_no')
        || lower.includes('missing groups')
        || lower.includes('missing csv hourly')
        || lower.includes('missing eod')
        || lower.includes('error')
      ) {
        return 'error';
      }
      if (
        lower.includes('no discrepancy found')
        || lower.includes(' match')
        || lower.includes('no per-transaction discrepancy found')
        || lower.includes('analyzed successfully')
      ) {
        return 'success';
      }
      return 'neutral';
    };
    const html = rawMessage
      .split('\n')
      .map((line) => {
        const tone = classifyLineTone(line);
        return `<span class="result-line result-line--${tone}">${escapeHtml(line)}</span>`;
      })
      .join('\n');
    ayalaResultText.innerHTML = html;
    ayalaResultText.classList.remove('result-output--error', 'result-output--success');
    ayalaResultText.classList.add(isError ? 'result-output--error' : 'result-output--success');
    ayalaResultModal.hidden = false;
  }

  function closeAyalaResultModalFn() {
    if (!ayalaResultModal) {
      return;
    }
    ayalaResultModal.hidden = true;
  }

  function resetAyalaReportOutput() {
    closeAyalaResultModalFn();
    if (ayalaReportOutput) {
      ayalaReportOutput.hidden = true;
      ayalaReportOutput.value = '';
      ayalaReportOutput.classList.remove('sql-output--error', 'sql-output--success');
    }
    if (copyAyalaReportBtn) {
      copyAyalaReportBtn.disabled = true;
    }
    if (copyAyalaReportRow) {
      copyAyalaReportRow.hidden = true;
    }
    if (downloadAyalaCheckerBtn) {
      downloadAyalaCheckerBtn.disabled = true;
    }
    if (ayalaDiscrepancyActions) {
      ayalaDiscrepancyActions.hidden = true;
    }
    latestAyalaTxnValidationReport = '';
    latestAyalaTotalValidationReport = '';
    latestAyalaCheckerData = null;
    if (ayalaStatus) {
      ayalaStatus.classList.remove('analyze-status--error', 'analyze-status--success');
    }
  }

  function setAyalaReportTone(isError) {
    if (ayalaReportOutput) {
      ayalaReportOutput.classList.remove('sql-output--error', 'sql-output--success');
      // Keep main textarea neutral. Detailed red/green is shown per-line in result modal.
    }
    if (ayalaStatus) {
      ayalaStatus.classList.remove('analyze-status--error', 'analyze-status--success');
      ayalaStatus.classList.add(isError ? 'analyze-status--error' : 'analyze-status--success');
    }
  }

  if (checkMissingOrBtn) {
    checkMissingOrBtn.addEventListener('click', openMissingOrModal);
  }

  if (openGenerateFrecnoBtn) {
    openGenerateFrecnoBtn.addEventListener('click', openFrecnoModal);
  }

  if (openFrecnoListSqlBtn) {
    openFrecnoListSqlBtn.addEventListener('click', openFrecnoListSqlModal);
  }

  if (openGenerateAllScriptBtn) {
    openGenerateAllScriptBtn.addEventListener('click', openAllScriptModal);
  }

  if (openXmlRenameBtn) {
    openXmlRenameBtn.addEventListener('click', openXmlRenameModal);
  }

  if (openReportDocxBtn) {
    openReportDocxBtn.addEventListener('click', openReportDocxModal);
  }

  if (closeMissingOrModal) {
    closeMissingOrModal.addEventListener('click', closeMissingOrModalFn);
  }

  if (closeFrecnoModal) {
    closeFrecnoModal.addEventListener('click', closeFrecnoModalFn);
  }

  if (closeFrecnoListSqlModal) {
    closeFrecnoListSqlModal.addEventListener('click', closeFrecnoListSqlModalFn);
  }

  if (closeAllScriptModal) {
    closeAllScriptModal.addEventListener('click', closeAllScriptModalFn);
  }

  if (closeXmlRenameModal) {
    closeXmlRenameModal.addEventListener('click', closeXmlRenameModalFn);
  }

  if (closeReportDocxModal) {
    closeReportDocxModal.addEventListener('click', closeReportDocxModalFn);
  }

  if (generateReportDocxBtn) {
    generateReportDocxBtn.addEventListener('click', async () => {
      const hasActivitiesEditor = Boolean(reportActivitiesEditor);
      const originalActivitiesHtml = hasActivitiesEditor ? reportActivitiesEditor.innerHTML : '';
      await convertActivitiesEditorExternalImages({ replaceOnFail: false });
      const payload = {
        date: formatReportDateMmDdYyyy(reportDateInput ? reportDateInput.value.trim() : ''),
        jof: reportJofInput ? reportJofInput.value.trim() : '',
        invoice: reportInvoiceInput ? reportInvoiceInput.value.trim() : '',
        client: reportClientInput ? reportClientInput.value.trim() : '',
        tsAssigned: reportTsAssignedInput ? reportTsAssignedInput.value.trim() : '',
        title: reportTitleInput ? reportTitleInput.value.trim() : 'Status Report',
        concern: reportConcernInput ? reportConcernInput.value.trim() : '',
        activities: getActivitiesEditorText(),
        activitiesImages: getActivitiesEditorImages(),
        activitiesBlocks: getActivitiesBlocksFromEditor(),
        rootCause: reportRootCauseInput ? reportRootCauseInput.value.trim() : '',
        preventiveAction: reportPreventiveInput ? reportPreventiveInput.value.trim() : '',
        nextSteps: reportNextStepsInput ? reportNextStepsInput.value.trim() : '',
        currentStatus: reportCurrentStatusInput ? reportCurrentStatusInput.value.trim() : '',
        body: getActivitiesEditorText(),
      };
      if (generateReportDocxBtn) generateReportDocxBtn.disabled = true;
      if (clearReportDocxBtn) clearReportDocxBtn.disabled = true;
      try {
        if (reportDocxStatus) {
          reportDocxStatus.hidden = false;
          reportDocxStatus.textContent = 'Generating DOCX...';
        }
        let templateBuffer;
        if (reportDocxTemplateInput && reportDocxTemplateInput.files && reportDocxTemplateInput.files[0]) {
          templateBuffer = await reportDocxTemplateInput.files[0].arrayBuffer();
        } else {
          templateBuffer = await readTemplateFromFilesFolder();
        }
        const blob = await buildReportDocxBlobFromTemplate(templateBuffer, payload);
        const usedTemplateName = REPORT_DOCX_TEMPLATE_PATH;
        const safeTitle = String(payload.title || 'Status Report')
          .replace(/[\\/:*?"<>|]/g, ' ')
          .replace(/\s+/g, ' ')
          .trim() || 'Status Report';
        const fileName = `${safeTitle}.docx`;
        downloadDocxBlob(blob, fileName);
        if (reportDocxOutput) {
          reportDocxOutput.value = [
            `Generated: ${fileName}`,
            `Template: ${usedTemplateName}`,
            `Date: ${payload.date || '-'}`,
            `JOF: ${payload.jof || '-'}`,
            `Invoice: ${payload.invoice || '-'}`,
            `Client: ${payload.client || '-'}`,
            `TS Assigned: ${payload.tsAssigned || '-'}`,
            '',
            'Section Mapping:',
            `Concern/s: ${payload.concern || payload.title || '-'}`,
            `Activities of TS: ${payload.activities ? 'filled from input' : '-'}`,
            `Root Cause: ${payload.rootCause ? 'filled from input' : 'kept template or parsed from sections'}`,
            `Preventive Action: ${payload.preventiveAction ? 'filled from input' : 'kept template or parsed from sections'}`,
            `Next Steps: ${payload.nextSteps ? 'filled from input' : 'kept template or parsed from sections'}`,
            `Current Status: ${payload.currentStatus ? 'filled from input' : 'kept template or parsed from sections'}`,
          ].join('\n');
        }
        if (reportDocxStatus) {
          reportDocxStatus.textContent = 'DOCX generated and downloaded.';
        }
      } catch (error) {
        if (reportDocxStatus) {
          reportDocxStatus.hidden = false;
          reportDocxStatus.textContent = 'Generate failed. Choose template manually or run via local server.';
        }
      } finally {
        if (hasActivitiesEditor && reportActivitiesEditor) {
          reportActivitiesEditor.innerHTML = originalActivitiesHtml;
        }
        if (generateReportDocxBtn) generateReportDocxBtn.disabled = false;
        if (clearReportDocxBtn) clearReportDocxBtn.disabled = false;
      }
    });
  }

  if (clearReportDocxBtn) {
    clearReportDocxBtn.addEventListener('click', () => {
      closeActivitiesEditorDialog();
      if (reportDocxTemplateInput) reportDocxTemplateInput.value = '';
      if (reportDateInput) reportDateInput.value = '';
      if (reportJofInput) reportJofInput.value = '';
      if (reportInvoiceInput) reportInvoiceInput.value = '';
      if (reportClientInput) reportClientInput.value = '';
      if (reportTsAssignedInput) reportTsAssignedInput.value = '';
      if (reportTitleInput) reportTitleInput.value = '';
      if (reportConcernInput) reportConcernInput.value = 'Ticket no. ';
      if (reportActivitiesEditor) reportActivitiesEditor.innerHTML = '';
      if (reportRootCauseInput) reportRootCauseInput.value = '';
      if (reportPreventiveInput) reportPreventiveInput.value = '';
      if (reportNextStepsInput) reportNextStepsInput.value = '';
      if (reportCurrentStatusInput) reportCurrentStatusInput.value = '';
      if (activitiesEditorState) activitiesEditorState.textContent = 'Click the button to input Activities of TS.';
      if (reportDocxOutput) reportDocxOutput.value = '';
      if (reportDocxStatus) {
        reportDocxStatus.hidden = false;
        reportDocxStatus.textContent = 'Cleared.';
      }
    });
  }

  if (downloadReportTemplateBtn) {
    downloadReportTemplateBtn.addEventListener('click', async () => {
      try {
        await downloadReportTemplateAndPromptSelect();
      } catch (error) {
        if (reportDocxStatus) {
          reportDocxStatus.hidden = false;
          reportDocxStatus.textContent = 'Template download failed.';
        }
      }
    });
  }

  if (activitiesBulletBtn) {
    activitiesBulletBtn.addEventListener('click', () => {
      applyActivitiesBulletsByLine();
    });
  }

  if (openActivitiesEditorBtn) {
    openActivitiesEditorBtn.addEventListener('click', () => {
      openActivitiesEditorDialog();
    });
  }

  if (closeActivitiesEditorBtn) {
    closeActivitiesEditorBtn.addEventListener('click', () => {
      closeActivitiesEditorDialog();
    });
  }

  if (okActivitiesEditorBtn) {
    okActivitiesEditorBtn.addEventListener('click', () => {
      closeActivitiesEditorDialog();
      if (activitiesEditorState) {
        const hasActivities = Boolean(String(reportActivitiesEditor && reportActivitiesEditor.innerText ? reportActivitiesEditor.innerText : '').trim());
        activitiesEditorState.textContent = hasActivities
          ? 'Activities of TS saved.'
          : 'Activities of TS is empty.';
      }
      if (reportDocxStatus) {
        reportDocxStatus.hidden = false;
        reportDocxStatus.textContent = 'Activities of TS saved.';
      }
    });
  }

  if (activitiesEditorDialog) {
    activitiesEditorDialog.addEventListener('click', (event) => {
      if (event.target === activitiesEditorDialog) {
        closeActivitiesEditorDialog();
      }
    });
  }

  if (generatePfpBtn) {
    generatePfpBtn.addEventListener('click', async () => {
      const promptValue = pfpPromptInput ? pfpPromptInput.value.trim() : '';
      const file = pfpFileInput && pfpFileInput.files ? pfpFileInput.files[0] : null;
      if (!promptValue && !file) {
        if (pfpStatus) pfpStatus.textContent = 'Please input prompt or choose photo first.';
        return;
      }
      if (generatePfpBtn) generatePfpBtn.disabled = true;
      if (downloadPfpBtn) downloadPfpBtn.disabled = true;
      try {
        if (pfpStatus) pfpStatus.textContent = promptValue ? 'Generating online from prompt...' : 'Generating from uploaded photo...';
        const result = promptValue ? await generateImageFromPrompt(promptValue) : { img: await readImageFile(file), corsSafe: true };
        const img = result.img;
        latestPfpImage = img;
        drawImageCoverToCanvas(img);
        latestPfpCanvasExportSafe = Boolean(result.corsSafe);

        if (latestPfpCanvasExportSafe) {
          await overlayAllianceLogoOnCanvas();
          latestPfpCanvasExportSafe = testCanvasExportSafety();
        }

        if (pfpStatus) {
          pfpStatus.textContent = latestPfpCanvasExportSafe
            ? 'Generated. Ready to download.'
            : 'Generated online. Direct image download mode enabled.';
        }
        if (downloadPfpBtn) downloadPfpBtn.disabled = false;
      } catch (error) {
        if (promptValue && latestPfpRemoteUrl) {
          const opened = openPfpRemoteUrlInNewTab();
          if (pfpStatus) {
            pfpStatus.textContent = opened
              ? 'Browser blocked in-app generation. Opened generator in new tab, save the image there.'
              : 'Generate failed. Please try again.';
          }
        } else if (pfpStatus) {
          pfpStatus.textContent = 'Generate failed. Check internet for prompt mode or use PNG/JPG/WebP upload.';
        }
      } finally {
        if (generatePfpBtn) generatePfpBtn.disabled = false;
      }
    });
  }

  if (generatePfpStableBtn) {
    generatePfpStableBtn.addEventListener('click', async () => {
      const apiKey = pfpApiKeyInput ? pfpApiKeyInput.value.trim() : '';
      const promptValue = pfpPromptInput ? pfpPromptInput.value.trim() : '';
      if (!apiKey) {
        if (pfpStatus) pfpStatus.textContent = 'Please input OpenAI API key.';
        return;
      }
      if (!promptValue) {
        if (pfpStatus) pfpStatus.textContent = 'Please input prompt first.';
        return;
      }
      if (generatePfpStableBtn) generatePfpStableBtn.disabled = true;
      if (generatePfpBtn) generatePfpBtn.disabled = true;
      if (downloadPfpBtn) downloadPfpBtn.disabled = true;
      try {
        if (pfpStatus) pfpStatus.textContent = 'Generating stable image via OpenAI...';
        const result = await generateImageFromOpenAI(apiKey, promptValue);
        latestPfpImage = result.img;
        drawImageCoverToCanvas(result.img);
        latestPfpCanvasExportSafe = Boolean(result.corsSafe);
        if (latestPfpCanvasExportSafe) {
          await overlayAllianceLogoOnCanvas();
          latestPfpCanvasExportSafe = testCanvasExportSafety();
        }
        if (downloadPfpBtn) downloadPfpBtn.disabled = false;
        if (pfpStatus) {
          pfpStatus.textContent = latestPfpCanvasExportSafe
            ? 'Stable generation done. Ready to download.'
            : 'Stable generation done. Open/download from new tab mode.';
        }
      } catch (error) {
        if (pfpStatus) pfpStatus.textContent = 'Stable generation failed. Check API key/billing/prompt.';
      } finally {
        if (generatePfpStableBtn) generatePfpStableBtn.disabled = false;
        if (generatePfpBtn) generatePfpBtn.disabled = false;
      }
    });
  }

  if (downloadPfpBtn) {
    downloadPfpBtn.addEventListener('click', () => {
      if (!latestPfpImage) {
        if (pfpStatus) pfpStatus.textContent = 'Generate first before download.';
        return;
      }
      if (latestPfpCanvasExportSafe) {
        downloadPfpCanvasPng();
        if (pfpStatus) pfpStatus.textContent = 'Downloaded: instagram_profile_1080.png';
        return;
      }
      if (latestPfpRemoteUrl) {
        openPfpRemoteUrlInNewTab();
        if (pfpStatus) pfpStatus.textContent = 'Opened generated image. Save it from the new tab.';
        return;
      }
      if (pfpStatus) pfpStatus.textContent = 'Download not available for this generated image.';
    });
  }

  if (clearPfpBtn) {
    clearPfpBtn.addEventListener('click', () => {
      if (pfpPromptInput) pfpPromptInput.value = '';
      if (pfpApiKeyInput) pfpApiKeyInput.value = '';
      if (pfpFileInput) pfpFileInput.value = '';
      latestPfpImage = null;
      latestPfpRemoteUrl = '';
      latestPfpCanvasExportSafe = true;
      if (latestPfpBlobUrl) {
        URL.revokeObjectURL(latestPfpBlobUrl);
        latestPfpBlobUrl = '';
      }
      drawPfpPlaceholder();
      if (downloadPfpBtn) downloadPfpBtn.disabled = true;
      if (pfpStatus) pfpStatus.textContent = 'Cleared.';
    });
  }

  if (reportActivitiesEditor) {
    reportActivitiesEditor.addEventListener('paste', (event) => {
      handleActivitiesPaste(event).catch(() => {
        // Keep editor usable even if image paste fails.
      });
    });
    reportActivitiesEditor.addEventListener('click', (event) => {
      const clickedImg = event.target && event.target.closest ? event.target.closest('img') : null;
      Array.from(reportActivitiesEditor.querySelectorAll('img.activities-inline-image--selected'))
        .forEach((img) => img.classList.remove('activities-inline-image--selected'));
      if (clickedImg && reportActivitiesEditor.contains(clickedImg)) {
        clickedImg.classList.add('activities-inline-image--selected');
      }
    });
    reportActivitiesEditor.addEventListener('keydown', (event) => {
      const key = String(event.key || '').toLowerCase();
      if (key !== 'backspace' && key !== 'delete') return;
      const selectedImg = reportActivitiesEditor.querySelector('img.activities-inline-image--selected');
      if (!selectedImg) return;
      event.preventDefault();
      const parent = selectedImg.parentNode;
      if (parent && parent !== reportActivitiesEditor && parent.childNodes.length === 1) {
        parent.remove();
      } else {
        selectedImg.remove();
      }
    });
  }

  if (processXmlRenameBtn) {
    processXmlRenameBtn.addEventListener('click', async () => {
      const files = xmlRenameFileInput && xmlRenameFileInput.files ? Array.from(xmlRenameFileInput.files) : [];
      const tenantIdValue = xmlTenantIdInput ? xmlTenantIdInput.value.trim() : '';
      const keyValue = xmlKeyInput ? xmlKeyInput.value.trim() : '';

      if (xmlRenameOutput) {
        xmlRenameOutput.value = '';
      }
      if (xmlRenameStatus) {
        xmlRenameStatus.hidden = false;
        xmlRenameStatus.textContent = 'Processing XML files...';
      }

      if (files.length === 0) {
        if (xmlRenameStatus) {
          xmlRenameStatus.textContent = 'Please choose XML files first.';
        }
        return;
      }
      if (!tenantIdValue) {
        if (xmlRenameStatus) {
          xmlRenameStatus.textContent = 'Please input Tenant ID.';
        }
        return;
      }
      if (!keyValue) {
        if (xmlRenameStatus) {
          xmlRenameStatus.textContent = 'Please input Key.';
        }
        return;
      }

      if (processXmlRenameBtn) {
        processXmlRenameBtn.disabled = true;
      }
      if (clearXmlRenameBtn) {
        clearXmlRenameBtn.disabled = true;
      }

      try {
        const logs = [];
        for (let i = 0; i < files.length; i += 1) {
          const file = files[i];
          const fileText = await readTextFile(file);
          let updated = fileText;
          updated = updateXmlTagValue(updated, 'tenantid', tenantIdValue);
          updated = updateXmlTagValue(updated, 'key', keyValue);
          const newName = buildRenamedXmlFileName(file.name, tenantIdValue);
          downloadXmlBlob(updated, newName);
          logs.push(`${file.name}  ->  ${newName}`);
          if (xmlRenameStatus) {
            xmlRenameStatus.textContent = `Processing XML files... (${i + 1}/${files.length})`;
          }
          await new Promise((resolve) => window.setTimeout(resolve, 0));
        }

        if (xmlRenameOutput) {
          xmlRenameOutput.value = logs.join('\n');
        }
        if (xmlRenameStatus) {
          xmlRenameStatus.textContent = `Done. ${files.length} XML file(s) processed and downloaded.`;
        }
      } catch (error) {
        if (xmlRenameStatus) {
          xmlRenameStatus.textContent = 'XML rename failed. Please check files and try again.';
        }
      } finally {
        if (processXmlRenameBtn) {
          processXmlRenameBtn.disabled = false;
        }
        if (clearXmlRenameBtn) {
          clearXmlRenameBtn.disabled = false;
        }
      }
    });
  }

  if (clearXmlRenameBtn) {
    clearXmlRenameBtn.addEventListener('click', () => {
      if (xmlRenameFileInput) {
        xmlRenameFileInput.value = '';
      }
      if (xmlTenantIdInput) {
        xmlTenantIdInput.value = '';
      }
      if (xmlKeyInput) {
        xmlKeyInput.value = '';
      }
      if (xmlRenameOutput) {
        xmlRenameOutput.value = '';
      }
      if (xmlRenameStatus) {
        xmlRenameStatus.hidden = false;
        xmlRenameStatus.textContent = 'Cleared.';
      }
    });
  }

  if (orFileInput) {
    orFileInput.addEventListener('change', () => {
      // Selecting a new file resets old output; processing happens on Upload click.
      resetMissingOutput();
    });
  }

  if (uploadOrFileBtn) {
    uploadOrFileBtn.addEventListener('click', () => {
      const file = orFileInput && orFileInput.files ? orFileInput.files[0] : null;
      resetMissingOutput();
      if (analyzeStatus) {
        analyzeStatus.hidden = false;
        analyzeStatus.textContent = 'Analyzing file, please wait...';
      }

      if (!file) {
        if (missingSummary && detectedCount && missingCount && outOfOrderCount && missingListOutput) {
          detectedCount.textContent = '0';
          missingCount.textContent = '0';
          if (duplicateCount) duplicateCount.textContent = '0';
          if (duplicateDetailMismatchCount) duplicateDetailMismatchCount.textContent = '0';
          if (rollbackCount) rollbackCount.textContent = '0';
          outOfOrderCount.textContent = '0';
          missingListOutput.value = 'Please choose a file first.';
          missingSummary.hidden = false;
        }
        if (analyzeStatus) {
          analyzeStatus.hidden = false;
          analyzeStatus.textContent = 'No file selected.';
        }
        return;
      }

      const fileName = file.name.toLowerCase();
      if (fileName.endsWith('.xlsx')) {
        const reader = new FileReader();
        reader.onload = () => {
          const buffer = reader.result;
          if (buffer instanceof ArrayBuffer) {
            processUploadedWorkbook(buffer);
          }
        };
        reader.readAsArrayBuffer(file);
        return;
      }

      const reader = new FileReader();
      reader.onload = () => {
        const text = typeof reader.result === 'string' ? reader.result : '';
        processUploadedText(text);
      };
      reader.readAsText(file);
    });
  }

  if (clearOrFileBtn) {
    clearOrFileBtn.addEventListener('click', () => {
      if (orFileInput) {
        orFileInput.value = '';
      }
      resetMissingOutput();
      if (analyzeStatus) {
        analyzeStatus.hidden = false;
        analyzeStatus.textContent = 'Cleared.';
      }
    });
  }

  if (generateSelectSqlBtn && sqlResultOutput) {
    generateSelectSqlBtn.addEventListener('click', () => {
      sqlResultOutput.value = buildSelectSqlFromMissing();
      sqlResultOutput.hidden = false;
      if (copyMissingSqlBtn) {
        copyMissingSqlBtn.disabled = false;
      }
      if (copyMissingSqlRow) {
        copyMissingSqlRow.hidden = false;
      }
    });
  }

  if (generateDeleteSqlBtn && sqlResultOutput) {
    generateDeleteSqlBtn.addEventListener('click', () => {
      sqlResultOutput.value = buildDeleteSqlFromMissing();
      sqlResultOutput.hidden = false;
      if (copyMissingSqlBtn) {
        copyMissingSqlBtn.disabled = false;
      }
      if (copyMissingSqlRow) {
        copyMissingSqlRow.hidden = false;
      }
    });
  }

  if (copyMissingSqlBtn) {
    copyMissingSqlBtn.disabled = true;
    copyMissingSqlBtn.addEventListener('click', async () => {
      if (!sqlResultOutput || !sqlResultOutput.value.trim()) {
        return;
      }

      try {
        if (navigator.clipboard && navigator.clipboard.writeText) {
          await navigator.clipboard.writeText(sqlResultOutput.value);
        } else {
          sqlResultOutput.focus();
          sqlResultOutput.select();
          document.execCommand('copy');
        }
        if (analyzeStatus) {
          analyzeStatus.hidden = false;
          analyzeStatus.textContent = 'SQL code copied.';
        }
      } catch (error) {
        if (analyzeStatus) {
          analyzeStatus.hidden = false;
          analyzeStatus.textContent = 'Copy failed. Please copy manually.';
        }
      }
    });
  }

  if (generateFrecnoListSelectBtn && frecnoListSqlOutput) {
    generateFrecnoListSelectBtn.addEventListener('click', () => {
      normalizeFrecnoInputToVertical();
      const rawText = frecnoListSqlInput ? frecnoListSqlInput.value : '';
      frecnoListSqlOutput.value = buildSelectSqlFromFrecnoList(rawText);
      if (copyFrecnoListSqlBtn) {
        copyFrecnoListSqlBtn.disabled = !frecnoListSqlOutput.value.trim() || frecnoListSqlOutput.value.startsWith('--');
      }
      if (copyFrecnoListSqlRow) {
        copyFrecnoListSqlRow.hidden = false;
      }
      if (frecnoListSqlStatus) {
        frecnoListSqlStatus.hidden = false;
        frecnoListSqlStatus.textContent = frecnoListSqlOutput.value.startsWith('--')
          ? 'Please enter valid frecno values.'
          : 'SELECT SQL generated.';
      }
    });
  }

  if (generateFrecnoListDeleteBtn && frecnoListSqlOutput) {
    generateFrecnoListDeleteBtn.addEventListener('click', () => {
      normalizeFrecnoInputToVertical();
      const rawText = frecnoListSqlInput ? frecnoListSqlInput.value : '';
      frecnoListSqlOutput.value = buildDeleteSqlFromFrecnoList(rawText);
      if (copyFrecnoListSqlBtn) {
        copyFrecnoListSqlBtn.disabled = !frecnoListSqlOutput.value.trim() || frecnoListSqlOutput.value.startsWith('--');
      }
      if (copyFrecnoListSqlRow) {
        copyFrecnoListSqlRow.hidden = false;
      }
      if (frecnoListSqlStatus) {
        frecnoListSqlStatus.hidden = false;
        frecnoListSqlStatus.textContent = frecnoListSqlOutput.value.startsWith('--')
          ? 'Please enter valid frecno values.'
          : 'DELETE SQL generated.';
      }
    });
  }

  if (copyFrecnoListSqlBtn) {
    copyFrecnoListSqlBtn.disabled = true;
    copyFrecnoListSqlBtn.addEventListener('click', async () => {
      if (!frecnoListSqlOutput || !frecnoListSqlOutput.value.trim() || frecnoListSqlOutput.value.startsWith('--')) {
        return;
      }
      try {
        if (navigator.clipboard && navigator.clipboard.writeText) {
          await navigator.clipboard.writeText(frecnoListSqlOutput.value);
        } else {
          frecnoListSqlOutput.focus();
          frecnoListSqlOutput.select();
          document.execCommand('copy');
        }
        if (frecnoListSqlStatus) {
          frecnoListSqlStatus.hidden = false;
          frecnoListSqlStatus.textContent = 'Code copied.';
        }
      } catch (error) {
        if (frecnoListSqlStatus) {
          frecnoListSqlStatus.hidden = false;
          frecnoListSqlStatus.textContent = 'Copy failed. Please copy manually.';
        }
      }
    });
  }

  if (clearFrecnoListSqlBtn) {
    clearFrecnoListSqlBtn.addEventListener('click', () => {
      if (frecnoListSqlInput) {
        frecnoListSqlInput.value = '';
      }
      if (frecnoListSqlOutput) {
        frecnoListSqlOutput.value = '';
      }
      if (copyFrecnoListSqlRow) {
        copyFrecnoListSqlRow.hidden = true;
      }
      if (copyFrecnoListSqlBtn) {
        copyFrecnoListSqlBtn.disabled = true;
      }
      if (frecnoListSqlStatus) {
        frecnoListSqlStatus.hidden = false;
        frecnoListSqlStatus.textContent = 'Ready.';
      }
      if (frecnoListSqlInput) {
        frecnoListSqlInput.focus();
      }
    });
  }

  if (frecnoListSqlInput) {
    frecnoListSqlInput.addEventListener('blur', () => {
      normalizeFrecnoInputToVertical();
    });
  }

  if (openAyalaCheckerBtn) {
    openAyalaCheckerBtn.addEventListener('click', openAyalaModal);
  }

  if (closeAyalaModal) {
    closeAyalaModal.addEventListener('click', closeAyalaModalFn);
  }

  if (closeAyalaResultModal) {
    closeAyalaResultModal.addEventListener('click', closeAyalaResultModalFn);
  }

  if (okAyalaResultBtn) {
    okAyalaResultBtn.addEventListener('click', closeAyalaResultModalFn);
  }

  if (ayalaCheckerTabBtn) {
    ayalaCheckerTabBtn.addEventListener('click', () => {
      setAyalaPanel('checker');
      if (ayalaStatus) {
        ayalaStatus.hidden = false;
        ayalaStatus.textContent = 'Ayala checker ready.';
      }
    });
  }

  if (ayalaGuidelinesTabBtn) {
    ayalaGuidelinesTabBtn.addEventListener('click', async () => {
      await downloadAyalaGuidelinesZip();
      setAyalaPanel('checker');
    });
  }

  async function runAyalaCrossValidation() {
    const hourlyFiles = hourlyCsvInput && hourlyCsvInput.files ? Array.from(hourlyCsvInput.files) : [];
    const eodFiles = eodCsvInput && eodCsvInput.files ? Array.from(eodCsvInput.files) : [];
    const currentSignature = getAyalaFilesSignature(hourlyFiles);
    resetAyalaReportOutput();
    if (ayalaStatus) {
      ayalaStatus.hidden = false;
      ayalaStatus.textContent = 'Analyzing files...';
    }

    if (hourlyFiles.length === 0 && eodFiles.length === 0) {
      if (ayalaStatus) {
        ayalaStatus.textContent = 'Please choose CSV Hourly files and EOD files.';
      }
      return;
    }
    if (hourlyFiles.length === 0) {
      if (ayalaStatus) {
        ayalaStatus.textContent = 'Please choose at least one CSV Hourly file.';
      }
      return;
    }
    if (eodFiles.length === 0) {
      if (ayalaStatus) {
        ayalaStatus.textContent = 'Please choose at least one EOD file.';
      }
      return;
    }
    if (!ayalaMergedHourlyData || !ayalaMergedHourlySignature || ayalaMergedHourlySignature !== currentSignature) {
      if (ayalaStatus) {
        ayalaStatus.textContent = 'Please click Merge first for the currently selected CSV Hourly files.';
      }
      return;
    }

    if (uploadAyalaBtn) {
      uploadAyalaBtn.disabled = true;
    }
    if (clearAyalaBtn) {
      clearAyalaBtn.disabled = true;
    }
    if (mergeAyalaBtn) {
      mergeAyalaBtn.disabled = true;
    }

    try {
      const hourlyData = ayalaMergedHourlyData;
      const eodData = await readAyalaEodFiles(eodFiles);

      if (hourlyData.parseError || eodData.parseError) {
        if (ayalaStatus) {
          ayalaStatus.textContent = hourlyData.parseError || eodData.parseError || 'Failed to parse files.';
        }
        setAyalaReportTone(true);
        return;
      }

      const hourlyRequired = ['CCCODE', 'MERCHANT_NAME', 'TRN_DATE', 'TER_NO', 'TRANSACTION_NO', 'GROSS_SLS'];
      const eodRequired = ['CCCODE', 'MERCHANT_NAME', 'TRN_DATE', 'TER_NO', 'STRANS', 'ETRANS', 'GROSS_SLS', 'NO_TRN'];
      const missingHourlyHeaders = checkRequiredAyalaHeaders(hourlyData.headers, hourlyRequired);
      const missingEodHeaders = checkRequiredAyalaHeaders(eodData.headers, eodRequired);

      if (missingHourlyHeaders.length > 0 || missingEodHeaders.length > 0) {
        const reportLines = ['AYALA VALIDATOR REPORT', ''];
        if (missingHourlyHeaders.length > 0) {
          reportLines.push(`Missing Hourly headers: ${missingHourlyHeaders.join(', ')}`);
        }
        if (missingEodHeaders.length > 0) {
          reportLines.push(`Missing EOD headers: ${missingEodHeaders.join(', ')}`);
        }
        if (ayalaReportOutput) {
          ayalaReportOutput.hidden = false;
          ayalaReportOutput.value = reportLines.join('\n');
        }
        if (copyAyalaReportBtn) {
          copyAyalaReportBtn.disabled = false;
        }
        if (copyAyalaReportRow) {
          copyAyalaReportRow.hidden = false;
        }
        if (ayalaStatus) {
          ayalaStatus.textContent = 'Validation failed: required columns missing.';
        }
        setAyalaReportTone(true);
        latestAyalaTotalValidationReport = reportLines.join('\n');
        latestAyalaTxnValidationReport = '';
        if (ayalaDiscrepancyActions) {
          ayalaDiscrepancyActions.hidden = false;
        }
        return;
      }

      if (ayalaStatus) {
        ayalaStatus.textContent = 'Finalizing validation report...';
      }
      await new Promise((resolve) => window.setTimeout(resolve, 0));

      const hourlySourceDateMap = buildAyalaSourceFileDateMap(hourlyData.records);
      const eodSourceDateMap = buildAyalaSourceFileDateMap(eodData.records);
      const hourlyByDate = groupAyalaRecordsByDate(hourlyData.records, hourlySourceDateMap);
      const eodByDate = groupAyalaRecordsByDate(eodData.records, eodSourceDateMap);
      const eodRawByDate = buildAyalaEodRawPairsByDate(eodData.rawPairsByFile, eodSourceDateMap);
      const dateKeys = Array.from(new Set([...hourlyByDate.keys(), ...eodByDate.keys()]))
        .sort((a, b) => String(a).localeCompare(String(b)));

      const dateGroups = [];
      const totalReportSections = [];
      const txnReportSections = [];
      const missingDateSections = [];

      dateKeys.forEach((dateKey) => {
        const hourlyRecordsForDate = hourlyByDate.get(dateKey) || [];
        const eodRecordsForDate = eodByDate.get(dateKey) || [];
        if (hourlyRecordsForDate.length === 0 || eodRecordsForDate.length === 0) {
          const missingPart = hourlyRecordsForDate.length === 0
            ? 'Missing CSV Hourly files for this date.'
            : 'Missing EOD file for this date.';
          missingDateSections.push(`[DATE ${dateKey}] ${missingPart}`);
          return;
        }

        const validationResultByDate = validateAyalaData(hourlyRecordsForDate, eodRecordsForDate);
        const mergedRawGrids = eodRawByDate.get(dateKey) || [];
        const eodRawPairsByDate = mergedRawGrids.length > 0
          ? mergeAyalaEodRawPairGrids(mergedRawGrids)
          : buildAyalaEodPairsByTerminal(eodRecordsForDate);
        dateGroups.push({
          dateKey,
          hourlyRecords: hourlyRecordsForDate,
          eodRecords: eodRecordsForDate,
          eodRawPairs: eodRawPairsByDate,
          validationResult: validationResultByDate,
        });

        totalReportSections.push(`[DATE ${dateKey}]`);
        totalReportSections.push(buildAyalaValidationReport(validationResultByDate));
        totalReportSections.push('');

        const txnText = buildAyalaTransactionDiscrepancyReport(hourlyRecordsForDate);
        txnReportSections.push(`[DATE ${dateKey}]`);
        txnReportSections.push(txnText);
        txnReportSections.push('');
      });

      if (missingDateSections.length > 0) {
        totalReportSections.push('[DATE MISMATCH]');
        totalReportSections.push(...missingDateSections);
        totalReportSections.push('');
      }

      if (dateGroups.length === 0) {
        if (ayalaStatus) {
          ayalaStatus.textContent = 'No matching date groups found between CSV Hourly and EOD filenames.';
        }
        setAyalaReportTone(true);
        return;
      }

      const mergedValidation = dateGroups.length === 1
        ? dateGroups[0].validationResult
        : validateAyalaData(hourlyData.records, eodData.records);

      const totalReportText = totalReportSections.join('\n').trim();
      const txnReportText = txnReportSections.join('\n').trim() || 'No per-transaction discrepancy found.';
      if (ayalaReportOutput) {
        ayalaReportOutput.hidden = false;
        ayalaReportOutput.value = totalReportText;
      }
      if (copyAyalaReportBtn) {
        copyAyalaReportBtn.disabled = false;
      }
      if (copyAyalaReportRow) {
        copyAyalaReportRow.hidden = false;
      }
      latestAyalaCheckerData = {
        dateGroups,
        hourlyRecords: hourlyData.records,
        eodRecords: eodData.records,
        trnDate: mergedValidation.eodSummary.trnDate || '',
        terNo: mergedValidation.eodSummary.terNo || '',
        terNos: mergedValidation.eodSummary.terNos || [],
      };
      latestAyalaTxnValidationReport = txnReportText;
      latestAyalaTotalValidationReport = totalReportText;
      if (downloadAyalaCheckerBtn) {
        downloadAyalaCheckerBtn.disabled = false;
      }
      if (ayalaDiscrepancyActions) {
        ayalaDiscrepancyActions.hidden = false;
      }

      const hasIssueLine = /Cross Validation,\s*Discrepancy\b/i.test(totalReportText)
        || /Missing groups in (EOD|Hourly)\./i.test(totalReportText)
        || /Duplicate TRANSACTION_NO detected in CSV Hourly\./i.test(totalReportText);
      const hasIssues = hasIssueLine || missingDateSections.length > 0;
      if (ayalaStatus) {
        ayalaStatus.textContent = hasIssues
          ? 'Validation complete. Discrepancies found. Choose a validation result button to view details.'
          : 'Validation complete. No discrepancies found. Choose a validation result button to view details.';
      }
      setAyalaReportTone(hasIssues);

      if (typeof XLSX === 'undefined' && ayalaStatus) {
        ayalaStatus.hidden = false;
        ayalaStatus.textContent = 'Validation complete, but XLSX download is unavailable (XLSX library not loaded).';
      }
    } catch (error) {
      if (ayalaStatus) {
        ayalaStatus.textContent = 'Validation failed. Please check your files and try again.';
      }
      setAyalaReportTone(true);
    } finally {
      if (uploadAyalaBtn) {
        uploadAyalaBtn.disabled = false;
      }
      if (clearAyalaBtn) {
        clearAyalaBtn.disabled = false;
      }
      if (mergeAyalaBtn) {
        mergeAyalaBtn.disabled = false;
      }
    }
  }

  if (uploadAyalaBtn) {
    uploadAyalaBtn.addEventListener('click', async () => {
      await runAyalaCrossValidation();
    });
  }

  if (mergeAyalaBtn) {
    mergeAyalaBtn.addEventListener('click', async () => {
      const hourlyFiles = hourlyCsvInput && hourlyCsvInput.files ? Array.from(hourlyCsvInput.files) : [];
      const currentSignature = getAyalaFilesSignature(hourlyFiles);
      resetAyalaReportOutput();

      if (hourlyFiles.length === 0) {
        if (ayalaStatus) {
          ayalaStatus.hidden = false;
          ayalaStatus.textContent = 'Please choose CSV Hourly files first.';
        }
        return;
      }

      if (ayalaStatus) {
        ayalaStatus.hidden = false;
        ayalaStatus.textContent = 'Merging hourly files...';
      }
      if (mergeAyalaBtn) {
        mergeAyalaBtn.disabled = true;
      }
      if (uploadAyalaBtn) {
        uploadAyalaBtn.disabled = true;
      }
      if (clearAyalaBtn) {
        clearAyalaBtn.disabled = true;
      }

      try {
        const hourlyData = await readAyalaHourlyFiles(hourlyFiles);
        if (hourlyData.parseError) {
          if (ayalaStatus) {
            ayalaStatus.textContent = `Hourly analysis failed: ${hourlyData.parseError}`;
          }
          resetAyalaMergeState();
          setAyalaHourlyFirstMode(false);
          return;
        }
        ayalaMergedHourlyData = hourlyData;
        ayalaMergedHourlySignature = currentSignature;
        ayalaHourlyAnalyzed = true;
        setAyalaHourlyFirstMode(true);

        const mergedText = buildAyalaHourlyAnalysisReport(hourlyData.records);
        if (ayalaReportOutput) {
          ayalaReportOutput.hidden = false;
          ayalaReportOutput.value = mergedText;
        }
        if (copyAyalaReportBtn) {
          copyAyalaReportBtn.disabled = false;
        }
        if (copyAyalaReportRow) {
          copyAyalaReportRow.hidden = false;
        }
        if (ayalaStatus) {
          ayalaStatus.hidden = false;
          ayalaStatus.textContent = `Hourly analysis complete: ${hourlyData.records.length} rows processed. You can now choose EOD file for cross validation.`;
        }
        setAyalaReportTone(false);
      } catch (error) {
        if (ayalaStatus) {
          ayalaStatus.hidden = false;
          ayalaStatus.textContent = 'Hourly analysis failed. Please check files and try again.';
        }
        setAyalaReportTone(true);
      } finally {
        if (mergeAyalaBtn) {
          mergeAyalaBtn.disabled = false;
        }
        if (uploadAyalaBtn) {
          uploadAyalaBtn.disabled = false;
        }
        if (clearAyalaBtn) {
          clearAyalaBtn.disabled = false;
        }
      }
    });
  }

  if (clearAyalaBtn) {
    clearAyalaBtn.addEventListener('click', () => {
      if (hourlyCsvInput) {
        hourlyCsvInput.value = '';
      }
      if (eodCsvInput) {
        eodCsvInput.value = '';
      }
      resetAyalaMergeState();
      setAyalaHourlyFirstMode(false);
      resetAyalaReportOutput();
      if (ayalaStatus) {
        ayalaStatus.hidden = false;
        ayalaStatus.textContent = 'Ayala files cleared. Step 1: choose CSV hourly files then click Analyze Hourly.';
      }
    });
  }

  if (copyAyalaReportBtn) {
    copyAyalaReportBtn.disabled = true;
    copyAyalaReportBtn.addEventListener('click', async () => {
      if (!ayalaReportOutput || !ayalaReportOutput.value.trim()) {
        return;
      }

      try {
        if (navigator.clipboard && navigator.clipboard.writeText) {
          await navigator.clipboard.writeText(ayalaReportOutput.value);
        } else {
          ayalaReportOutput.focus();
          ayalaReportOutput.select();
          document.execCommand('copy');
        }
        if (ayalaStatus) {
          ayalaStatus.hidden = false;
          ayalaStatus.textContent = 'Report copied.';
        }
      } catch (error) {
        if (ayalaStatus) {
          ayalaStatus.hidden = false;
          ayalaStatus.textContent = 'Copy failed. Please copy manually.';
        }
      }
    });
  }

  if (downloadAyalaCheckerBtn) {
    downloadAyalaCheckerBtn.disabled = true;
    downloadAyalaCheckerBtn.addEventListener('click', () => {
      if (!latestAyalaCheckerData) {
        if (ayalaStatus) {
          ayalaStatus.hidden = false;
          ayalaStatus.textContent = 'No validation data yet. Please click Upload first.';
        }
        return;
      }
      if (typeof XLSX === 'undefined') {
        if (ayalaStatus) {
          ayalaStatus.hidden = false;
          ayalaStatus.textContent = 'XLSX download is unavailable (XLSX library not loaded).';
        }
        return;
      }

      const dateGroups = Array.isArray(latestAyalaCheckerData.dateGroups)
        ? latestAyalaCheckerData.dateGroups
        : [];
      const toAyalaFileToken = (value, fallback = 'MERCHANT') => {
        const token = String(value || '')
          .replace(/[\\/:*?"<>|]/g, ' ')
          .replace(/\s+/g, ' ')
          .trim();
        return token || fallback;
      };
      const getAyalaMerchantFromGroup = (group) => {
        const fromEod = Array.isArray(group && group.eodRecords)
          ? group.eodRecords.map((record) => String(record && record.MERCHANT_NAME || '').trim()).find(Boolean)
          : '';
        if (fromEod) return fromEod;
        const fromHourly = Array.isArray(group && group.hourlyRecords)
          ? group.hourlyRecords.map((record) => String(record && record.MERCHANT_NAME || '').trim()).find(Boolean)
          : '';
        return fromHourly || '';
      };
      let checkerWorkbook;
      let fileName;
      if (dateGroups.length > 1) {
        checkerWorkbook = XLSX.utils.book_new();
        dateGroups.forEach((group, index) => {
          const wbByDate = buildAyalaCheckerWorkbook(
            group.hourlyRecords || [],
            group.eodRecords || [],
            group.eodRawPairs || []
          );
          const safeDate = String(group.dateKey || `DATE${index + 1}`).replace(/[^0-9A-Za-z_-]/g, '_');
          const checkerSheet = wbByDate.Sheets.Checker;
          const txnSheet = wbByDate.Sheets['Txn Validation'];
          if (checkerSheet) {
            XLSX.utils.book_append_sheet(checkerWorkbook, checkerSheet, `Checker_${safeDate}`.slice(0, 31));
          }
          if (txnSheet) {
            XLSX.utils.book_append_sheet(checkerWorkbook, txnSheet, `Txn_${safeDate}`.slice(0, 31));
          }
        });
        const merchantSet = new Set(
          dateGroups
            .map((group) => getAyalaMerchantFromGroup(group))
            .filter(Boolean)
        );
        const merchantName = merchantSet.size === 1
          ? Array.from(merchantSet)[0]
          : 'MULTI MERCHANT';
        fileName = `Checker ${toAyalaFileToken(merchantName)}.xlsx`;
      } else {
        const single = dateGroups.length === 1 ? dateGroups[0] : latestAyalaCheckerData;
        checkerWorkbook = buildAyalaCheckerWorkbook(
          single.hourlyRecords || [],
          single.eodRecords || [],
          single.eodRawPairs || []
        );
        const merchantName = getAyalaMerchantFromGroup(single) || latestAyalaCheckerData.merchantName || 'MERCHANT';
        fileName = `Checker ${toAyalaFileToken(merchantName)}.xlsx`;
      }
      try {
        XLSX.writeFile(checkerWorkbook, fileName, { cellStyles: true });
        if (ayalaStatus) {
          ayalaStatus.hidden = false;
          ayalaStatus.textContent = 'Checker XLSX downloaded successfully.';
        }
      } catch (styleError) {
        try {
          XLSX.writeFile(checkerWorkbook, fileName);
          if (ayalaStatus) {
            ayalaStatus.hidden = false;
            ayalaStatus.textContent = 'Checker XLSX downloaded (style fallback mode).';
          }
        } catch (downloadError) {
          if (ayalaStatus) {
            ayalaStatus.hidden = false;
            ayalaStatus.textContent = 'Download failed. Please try again.';
          }
        }
      }
    });
  }

  if (ayalaTxnDiscrepancyBtn) {
    ayalaTxnDiscrepancyBtn.addEventListener('click', () => {
      if (!latestAyalaTxnValidationReport) {
        if (ayalaStatus) {
          ayalaStatus.hidden = false;
          ayalaStatus.textContent = 'No transaction validation result yet. Please click Upload first.';
        }
        return;
      }
      const hasIssues = !/No per-transaction discrepancy found\./.test(latestAyalaTxnValidationReport);
      openAyalaResultModal(latestAyalaTxnValidationReport, hasIssues);
    });
  }

  if (ayalaTotalDiscrepancyBtn) {
    ayalaTotalDiscrepancyBtn.addEventListener('click', () => {
      if (!latestAyalaTotalValidationReport) {
        if (ayalaStatus) {
          ayalaStatus.hidden = false;
          ayalaStatus.textContent = 'No hourly vs EOD validation result yet. Please click Upload first.';
        }
        return;
      }
      const hasIssues = /Cross Validation,\s*Discrepancy\b/i.test(latestAyalaTotalValidationReport)
        || /Missing groups in (EOD|Hourly)\./i.test(latestAyalaTotalValidationReport)
        || /Duplicate TRANSACTION_NO detected in CSV Hourly\./i.test(latestAyalaTotalValidationReport);
      openAyalaResultModal(latestAyalaTotalValidationReport, hasIssues);
    });
  }

  if (hourlyCsvInput) {
    hourlyCsvInput.addEventListener('change', () => {
      resetAyalaMergeState();
      setAyalaHourlyFirstMode(false);
      resetAyalaReportOutput();
      if (ayalaStatus) {
        ayalaStatus.hidden = false;
        ayalaStatus.textContent = 'Hourly files selected. Click Analyze Hourly first.';
      }
    });
  }

  if (eodCsvInput) {
    eodCsvInput.addEventListener('change', () => {
      const hasEod = eodCsvInput.files && eodCsvInput.files.length > 0;
      if (!hasEod) {
        return;
      }
      if (ayalaStatus) {
        ayalaStatus.hidden = false;
        ayalaStatus.textContent = 'EOD file(s) selected. Click Upload to run cross-validation.';
      }
    });
  }

  if (generateFrecnoBtn) {
    generateFrecnoBtn.addEventListener('click', async () => {
      const latestRaw = latestFrecnoInput ? latestFrecnoInput.value.trim() : '';
      const oldRaw = oldFrecnoInput ? oldFrecnoInput.value : '';

      if (!frecnoOutput || !frecnoStatus) {
        return;
      }

      if (frecnoProgress) {
        frecnoProgress.hidden = false;
        frecnoProgress.textContent = 'Analyzing inputs...';
      }
      frecnoStatus.textContent = 'Please wait while we analyze your frecno list.';
      frecnoOutput.value = '';
      hideFrecnoCopyButton();
      if (generateFrecnoBtn) {
        generateFrecnoBtn.disabled = true;
      }
      if (clearFrecnoBtn) {
        clearFrecnoBtn.disabled = true;
      }

      await new Promise((resolve) => window.setTimeout(resolve, 0));

      if (!isIntegerText(latestRaw)) {
        frecnoStatus.textContent = 'Please enter a valid latest frecno start value.';
        if (frecnoProgress) {
          frecnoProgress.hidden = true;
        }
        if (generateFrecnoBtn) {
          generateFrecnoBtn.disabled = false;
        }
        if (clearFrecnoBtn) {
          clearFrecnoBtn.disabled = false;
        }
        return;
      }

      const latestStart = Number.parseInt(latestRaw, 10);
      const oldFrecnos = parseFrecnoList(oldRaw);
      if (oldFrecnos.length === 0) {
        frecnoStatus.textContent = 'Please enter at least one old frecno.';
        if (frecnoProgress) {
          frecnoProgress.hidden = true;
        }
        if (generateFrecnoBtn) {
          generateFrecnoBtn.disabled = false;
        }
        if (clearFrecnoBtn) {
          clearFrecnoBtn.disabled = false;
        }
        return;
      }

      const duplicates = findDuplicateFrecnos(oldFrecnos);
      if (duplicates.length > 0) {
        frecnoStatus.textContent = `Duplicate old frecno found (${duplicates.length}). Please clean the list first.`;
        frecnoOutput.value = `Duplicates: ${duplicates.join(', ')}`;
        if (frecnoProgress) {
          frecnoProgress.hidden = true;
        }
        if (generateFrecnoBtn) {
          generateFrecnoBtn.disabled = false;
        }
        if (clearFrecnoBtn) {
          clearFrecnoBtn.disabled = false;
        }
        return;
      }

      if (frecnoProgress) {
        frecnoProgress.hidden = false;
        frecnoProgress.textContent = 'Finalizing SQL...';
      }

      const generatedSql = await buildFrecnoUpdateSql(oldFrecnos, latestStart);
      frecnoOutput.value = generatedSql;
      frecnoStatus.textContent = `Generated SQL mapping for ${oldFrecnos.length} frecno values.`;

      if (frecnoProgress) {
        frecnoProgress.hidden = true;
      }
      if (copyFrecnoCodeBtn) {
        copyFrecnoCodeBtn.disabled = false;
      }
      if (copyFrecnoCodeRow) {
        copyFrecnoCodeRow.hidden = false;
      }
      if (generateFrecnoBtn) {
        generateFrecnoBtn.disabled = false;
      }
      if (clearFrecnoBtn) {
        clearFrecnoBtn.disabled = false;
      }
    });
  }

  if (copyFrecnoCodeBtn) {
    hideFrecnoCopyButton();
    copyFrecnoCodeBtn.addEventListener('click', async () => {
      if (!frecnoOutput || !frecnoOutput.value.trim()) {
        return;
      }

      try {
        if (navigator.clipboard && navigator.clipboard.writeText) {
          await navigator.clipboard.writeText(frecnoOutput.value);
        } else {
          frecnoOutput.focus();
          frecnoOutput.select();
          document.execCommand('copy');
        }
        if (frecnoStatus) {
          frecnoStatus.textContent = 'Code copied.';
        }
      } catch (error) {
        if (frecnoStatus) {
          frecnoStatus.textContent = 'Copy failed. Please copy manually.';
        }
      }
    });
  }

  if (copyAllScriptBtn) {
    copyAllScriptBtn.addEventListener('click', async () => {
      const text = buildAllScriptPlainText();
      if (!text.trim()) {
        return;
      }
      try {
        if (navigator.clipboard && navigator.clipboard.writeText) {
          await navigator.clipboard.writeText(text);
        } else if (allScriptOutput) {
          const range = document.createRange();
          range.selectNodeContents(allScriptOutput);
          const selection = window.getSelection();
          if (selection) {
            selection.removeAllRanges();
            selection.addRange(range);
            document.execCommand('copy');
            selection.removeAllRanges();
          }
        }
      } catch (error) {
        // keep silent; user can still manually copy from modal.
      }
    });
  }

  if (clearFrecnoBtn) {
    clearFrecnoBtn.addEventListener('click', () => {
      if (generateFrecnoBtn) {
        generateFrecnoBtn.disabled = false;
      }
      if (clearFrecnoBtn) {
        clearFrecnoBtn.disabled = false;
      }
      if (oldFrecnoInput) {
        oldFrecnoInput.value = '';
      }
      if (latestFrecnoInput) {
        latestFrecnoInput.value = '';
      }
      if (frecnoOutput) {
        frecnoOutput.value = '';
      }
      hideFrecnoCopyButton();
      if (frecnoProgress) {
        frecnoProgress.hidden = true;
      }
      if (frecnoStatus) {
        frecnoStatus.textContent = 'Generated output will appear here.';
      }
    });
  }

  if (latestFrecnoInput) {
    latestFrecnoInput.addEventListener('input', hideFrecnoCopyButton);
  }

  if (oldFrecnoInput) {
    oldFrecnoInput.addEventListener('input', hideFrecnoCopyButton);
  }

  setAyalaHourlyFirstMode(false);

  links.forEach((link) => {
    link.addEventListener('click', (event) => {
      const viewName = link.dataset.viewLink;
      if (!viewName) {
        return;
      }

      event.preventDefault();
      switchViewWithTransition(viewName);
    });
  });

  const initial = window.location.hash.replace('#', '');
  const hasInitial = Array.from(views).some((view) => view.dataset.view === initial);
  setActiveView(hasInitial ? initial : 'home', false);
})();




