// engine.js — EHV Meeting Minutes DOCX engine
// Hosted on GitHub. Fetched at runtime by per-meeting artifacts.
// Public API (assigned to window): init(meeting), renderPreview(meeting), downloadDocx()

(function () {

  // ----------------------------------------------------------------
  // JSZip loader
  // ----------------------------------------------------------------
  function loadJSZip() {
    return new Promise(function (resolve, reject) {
      if (window.JSZip) return resolve();
      var s = document.createElement('script');
      s.src = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
      s.onload = resolve;
      s.onerror = function () { reject(new Error('JSZip load failed')); };
      document.head.appendChild(s);
    });
  }

  // ----------------------------------------------------------------
  // CSS
  // ----------------------------------------------------------------
  var style = document.createElement('style');
  style.textContent = [
    ':root{--bg:#fff;--text:#1a1a1a;--border:#ddd;--accent:#1a4480;--light-bg:#f7f7f7}',
    '@media(prefers-color-scheme:dark){:root{--bg:#1a1a1a;--text:#e0e0e0;--border:#444;--accent:#5b9bd5;--light-bg:#2a2a2a}}',
    '*{box-sizing:border-box;margin:0;padding:0}',
    'body{font-family:"Segoe UI",Arial,sans-serif;background:var(--bg);color:var(--text);padding:24px;max-width:800px;margin:0 auto}',
    '.header{text-align:center;margin-bottom:24px}',
    '.header img{max-width:250px;margin-bottom:12px}',
    'h1{font-size:20px;text-decoration:underline;margin-bottom:16px}',
    '.meta{margin-bottom:6px;font-size:14px}',
    '.meta strong{font-weight:600}',
    '.section{margin-top:20px}',
    '.section-title{font-weight:700;font-size:14px;margin-bottom:6px;border-bottom:1px solid var(--border);padding-bottom:4px}',
    '.section p{font-size:14px;line-height:1.6;text-align:justify;margin-bottom:8px}',
    '.section ul{margin-left:24px;margin-bottom:8px}',
    '.section li{font-size:14px;line-height:1.6;margin-bottom:4px}',
    '.download-bar{position:sticky;top:0;background:var(--accent);color:white;padding:12px 20px;border-radius:8px;display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;z-index:10}',
    '.download-bar button{background:white;color:var(--accent);border:none;padding:8px 20px;border-radius:4px;font-weight:600;cursor:pointer;font-size:14px}',
    '.download-bar button:hover{opacity:0.9}',
    '.download-bar .filename{font-size:14px;font-weight:500}'
  ].join('');
  document.head.appendChild(style);

  // ----------------------------------------------------------------
  // Internal state
  // ----------------------------------------------------------------
  var _meeting = null;
  var _logoB64 = '';
  var LOGO_URL = 'https://raw.githubusercontent.com/jgelhard/EHV-Power-Logo/main/base64.txt';

  // ----------------------------------------------------------------
  // XML helpers
  // ----------------------------------------------------------------
  function escXml(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function contentTypes() {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n' +
      '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n' +
      '  <Default Extension="xml" ContentType="application/xml"/>\n' +
      '  <Default Extension="jpeg" ContentType="image/jpeg"/>\n' +
      '  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\n' +
      '  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>\n' +
      '  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>\n' +
      '  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>\n' +
      '  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>\n' +
      '</Types>';
  }

  function rootRels() {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n' +
      '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>\n' +
      '</Relationships>';
  }

  function docRels() {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n' +
      '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\n' +
      '  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>\n' +
      '  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>\n' +
      '  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>\n' +
      '</Relationships>';
  }

  function headerRels() {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n' +
      '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.jpeg"/>\n' +
      '</Relationships>';
  }

  function headerXml() {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"\n' +
      '       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"\n' +
      '       xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"\n' +
      '       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"\n' +
      '       xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">\n' +
      '  <w:p><w:pPr><w:jc w:val="center"/></w:pPr>\n' +
      '    <w:r><w:rPr><w:noProof/></w:rPr>\n' +
      '      <w:drawing>\n' +
      '        <wp:inline distT="0" distB="0" distL="0" distR="0">\n' +
      '          <wp:extent cx="2466975" cy="633420"/>\n' +
      '          <wp:docPr id="1" name="Logo"/>\n' +
      '          <a:graphic>\n' +
      '            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">\n' +
      '              <pic:pic>\n' +
      '                <pic:nvPicPr><pic:cNvPr id="1" name="image1.jpeg"/><pic:cNvPicPr/></pic:nvPicPr>\n' +
      '                <pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>\n' +
      '                <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2466975" cy="633420"/></a:xfrm>\n' +
      '                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>\n' +
      '              </pic:pic>\n' +
      '            </a:graphicData>\n' +
      '          </a:graphic>\n' +
      '        </wp:inline>\n' +
      '      </w:drawing>\n' +
      '    </w:r>\n' +
      '  </w:p>\n' +
      '</w:hdr>';
  }

  function stylesXml() {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n' +
      '  <w:docDefaults>\n' +
      '    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>\n' +
      '    <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>\n' +
      '  </w:docDefaults>\n' +
      '  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>\n' +
      '  <w:style w:type="paragraph" w:styleId="Header"><w:name w:val="Header"/>\n' +
      '    <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>\n' +
      '  <w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/>\n' +
      '    <w:pPr><w:ind w:left="720"/></w:pPr></w:style>\n' +
      '</w:styles>';
  }

  function settingsXml() {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n' +
      '  <w:defaultTabStop w:val="720"/>\n' +
      '  <w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>\n' +
      '</w:settings>';
  }

  function numberingXml() {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n' +
      '  <w:abstractNum w:abstractNumId="0">\n' +
      '    <w:lvl w:ilvl="0">\n' +
      '      <w:start w:val="1"/>\n' +
      '      <w:numFmt w:val="bullet"/>\n' +
      '      <w:lvlText w:val="&#xF0B7;"/>\n' +
      '      <w:lvlJc w:val="left"/>\n' +
      '      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>\n' +
      '      <w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>\n' +
      '    </w:lvl>\n' +
      '  </w:abstractNum>\n' +
      '  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>\n' +
      '</w:numbering>';
  }

  // ----------------------------------------------------------------
  // DOCX paragraph builders
  // ----------------------------------------------------------------
  function pTitle(text) {
    return '<w:p><w:pPr><w:pStyle w:val="Header"/><w:jc w:val="center"/>\n' +
      '    <w:rPr><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/><w:u w:val="single"/></w:rPr></w:pPr>\n' +
      '    <w:r><w:rPr><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/><w:u w:val="single"/></w:rPr>\n' +
      '    <w:t>' + escXml(text) + '</w:t></w:r></w:p>';
  }

  function pLabelValue(label, value) {
    return '<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr>\n' +
      '    <w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>' + escXml(label) + '</w:t></w:r>\n' +
      '    <w:r><w:t xml:space="preserve"> ' + escXml(value) + '</w:t></w:r></w:p>';
  }

  function pSectionHead(text) {
    return '<w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:b/><w:bCs/></w:rPr></w:pPr>\n' +
      '    <w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>' + escXml(text) + '</w:t></w:r></w:p>';
  }

  function pBody(text) {
    return '<w:p><w:pPr><w:jc w:val="both"/></w:pPr>\n' +
      '    <w:r><w:t>' + escXml(text) + '</w:t></w:r></w:p>';
  }

  function pBullet(text) {
    return '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>\n' +
      '    <w:spacing w:after="0"/></w:pPr>\n' +
      '    <w:r><w:t>' + escXml(text) + '</w:t></w:r></w:p>';
  }

  function pEmpty() { return '<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr></w:p>'; }

  function buildDocXml(m) {
    var body = '';
    body += pEmpty();
    body += pTitle(m.title);
    body += pEmpty();
    body += pLabelValue('Date:', m.date);
    body += pEmpty();
    body += pLabelValue('Attendees:', m.attendees);
    body += pEmpty();
    body += pLabelValue('Project:', m.project);
    body += pEmpty();
    body += pEmpty();

    body += pSectionHead('Summary');
    body += pEmpty();
    if (m.summary.text) body += pBody(m.summary.text);
    if (m.summary.bullets) m.summary.bullets.forEach(function (b) { body += pBullet(b); });
    body += pEmpty();

    if (m.keyDecisions) {
      body += pSectionHead('Key Decisions');
      body += pEmpty();
      if (m.keyDecisions.text) body += pBody(m.keyDecisions.text);
      if (m.keyDecisions.bullets) m.keyDecisions.bullets.forEach(function (b) { body += pBullet(b); });
      body += pEmpty();
    }

    if (m.openIssues) {
      body += pSectionHead('Open Issues');
      body += pEmpty();
      if (m.openIssues.text) body += pBody(m.openIssues.text);
      if (m.openIssues.bullets) m.openIssues.bullets.forEach(function (b) { body += pBullet(b); });
      body += pEmpty();
    }

    if (m.actionItems) {
      body += pSectionHead('Action Items');
      body += pEmpty();
      if (m.actionItems.text) body += pBody(m.actionItems.text);
      if (m.actionItems.bullets) m.actionItems.bullets.forEach(function (b) { body += pBullet(b); });
      body += pEmpty();
    }

    if (m.personalActionItems && m.personalActionItems.length > 0) {
      m.personalActionItems.forEach(function (pai) {
        body += pSectionHead('Action Items for ' + escXml(pai.name));
        body += pEmpty();
        pai.items.forEach(function (item) { body += pBullet(item); });
        body += pEmpty();
      });
    }

    if (m.nextMeeting) {
      body += pSectionHead('Next Meeting');
      body += pEmpty();
      if (m.nextMeeting.text) body += pBody(m.nextMeeting.text);
      if (m.nextMeeting.bullets) m.nextMeeting.bullets.forEach(function (b) { body += pBullet(b); });
      body += pEmpty();
    }

    if (m.distributionList) {
      body += pSectionHead('Distribution List');
      body += pEmpty();
      if (m.distributionList.text) body += pBody(m.distributionList.text);
      if (m.distributionList.bullets) m.distributionList.bullets.forEach(function (b) { body += pBullet(b); });
      body += pEmpty();
    }

    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"\n' +
      '            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n' +
      '  <w:body>\n' +
      '    ' + body + '\n' +
      '    <w:sectPr>\n' +
      '      <w:headerReference w:type="first" r:id="rId4"/>\n' +
      '      <w:pgSz w:w="12240" w:h="15840"/>\n' +
      '      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="567" w:footer="567" w:gutter="0"/>\n' +
      '      <w:titlePg/>\n' +
      '    </w:sectPr>\n' +
      '  </w:body>\n' +
      '</w:document>';
  }

  // ----------------------------------------------------------------
  // DOCX generation (JSZip)
  // ----------------------------------------------------------------
  function b64toBlob(b64) {
    var bin = atob(b64);
    var arr = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
    return arr;
  }

  function sanitizeFilename(name) {
    return name.replace(/[<>:"/\\|?*]/g, '_').replace(/\s+/g, ' ').trim();
  }

  async function generateDocx(m) {
    var zip = new JSZip();
    zip.file('[Content_Types].xml', contentTypes());
    zip.folder('_rels').file('.rels', rootRels());
    var word = zip.folder('word');
    word.file('document.xml', buildDocXml(m));
    word.file('styles.xml', stylesXml());
    word.file('settings.xml', settingsXml());
    word.file('numbering.xml', numberingXml());
    word.file('header1.xml', headerXml());
    word.folder('_rels').file('document.xml.rels', docRels());
    word.folder('_rels').file('header1.xml.rels', headerRels());
    word.folder('media').file('image1.jpeg', b64toBlob(_logoB64), { binary: true });
    return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  }

  // ----------------------------------------------------------------
  // Public: downloadDocx — called by the rendered button
  // ----------------------------------------------------------------
  window.downloadDocx = async function () {
    if (!_meeting) return;
    var blob = await generateDocx(_meeting);
    var filename = sanitizeFilename(_meeting.title) + ' - Meeting Minutes - ' + sanitizeFilename(_meeting.date) + '.docx';
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  // ----------------------------------------------------------------
  // Public: renderPreview
  // ----------------------------------------------------------------
  window.renderPreview = function (m) {
    var app = document.getElementById('app');
    var html = '';

    html += '<div class="download-bar">' +
      '<span class="filename">' + sanitizeFilename(m.title) + ' - Meeting Minutes - ' + sanitizeFilename(m.date) + '.docx</span>' +
      '<button onclick="downloadDocx()">Download Word Document</button>' +
      '</div>';

    html += '<div class="header">' +
      (_logoB64 ? '<img src="data:image/jpeg;base64,' + _logoB64 + '" alt="EHV Power Logo"/>' : '') +
      '<h1>' + escXml(m.title) + '</h1>' +
      '</div>';

    html += '<div class="meta"><strong>Date:</strong> ' + escXml(m.date) + '</div>';
    html += '<div class="meta"><strong>Attendees:</strong> ' + escXml(m.attendees) + '</div>';
    html += '<div class="meta"><strong>Project:</strong> ' + escXml(m.project) + '</div>';

    function renderSection(title, data) {
      if (!data) return '';
      var s = '<div class="section"><div class="section-title">' + escXml(title) + '</div>';
      if (data.text) s += '<p>' + escXml(data.text) + '</p>';
      if (data.bullets && data.bullets.length) {
        s += '<ul>' + data.bullets.map(function (b) { return '<li>' + escXml(b) + '</li>'; }).join('') + '</ul>';
      }
      return s + '</div>';
    }

    html += renderSection('Summary', m.summary);
    html += renderSection('Key Decisions', m.keyDecisions);
    html += renderSection('Open Issues', m.openIssues);
    html += renderSection('Action Items', m.actionItems);

    if (m.personalActionItems && m.personalActionItems.length) {
      m.personalActionItems.forEach(function (pai) {
        html += renderSection('Action Items for ' + pai.name, { bullets: pai.items });
      });
    }

    html += renderSection('Next Meeting', m.nextMeeting);
    html += renderSection('Distribution List', m.distributionList);

    app.innerHTML = html;
  };

  // ----------------------------------------------------------------
  // Public: init — entry point called by per-meeting artifacts
  // ----------------------------------------------------------------
  window.init = async function (meeting) {
    _meeting = meeting;
    var app = document.getElementById('app');
    app.innerHTML = '<p style="text-align:center;padding:40px;color:#888;font-family:Arial,sans-serif">Loading...</p>';
    await loadJSZip();
    try {
      var resp = await fetch(LOGO_URL);
      if (!resp.ok) throw new Error('Logo fetch failed: ' + resp.status);
      _logoB64 = (await resp.text()).trim();
    } catch (e) {
      console.warn('Logo unavailable, rendering without it.', e);
      _logoB64 = '';
    }
    window.renderPreview(meeting);
  };

})();
