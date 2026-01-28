/**
 * Document Generation Service
 * Creates Word, Excel documents and HTML/PDF from AI responses
 * Uses browser-compatible libraries with LAZY LOADING for performance
 * Supports custom user branding (colors, logos, contact info)
 * 
 * Built by FreedomForged_AI
 */

import { loadBranding, UserBranding, getContactBlock, getFormattedAddress } from './brandingService';

export type DocumentType = 'word' | 'pdf' | 'excel' | 'powerpoint';

export type TemplateType = 
  | 'professional-report' 
  | 'meeting-summary' 
  | 'project-status' 
  | 'data-analysis' 
  | 'sales-pitch'
  | 'email-summary'
  | 'action-items'
  | 'custom';

// Lazy-loaded modules (reduces initial bundle by ~50MB)
let docxModule: typeof import('docx') | null = null;
let excelModule: typeof import('exceljs') | null = null;

async function loadDocx() {
  if (!docxModule) {
    docxModule = await import('docx');
  }
  return docxModule;
}

async function loadExcel() {
  if (!excelModule) {
    excelModule = await import('exceljs');
  }
  return excelModule;
}

export interface GeneratedDocument {
  blob: Blob;
  filename: string;
  mimeType: string;
}

export interface DocumentOptions {
  template?: TemplateType;
  title?: string;
  author?: string;
  includeCharts?: boolean;
  includeTableOfContents?: boolean;
  colorScheme?: 'default' | 'professional' | 'corporate' | 'creative';
  useCustomBranding?: boolean;  // Use user's saved branding
}

/**
 * Get colors - use custom branding if available, otherwise template colors
 */
function getColors(template: TemplateType, branding?: UserBranding): { primary: string; secondary: string; accent: string } {
  if (branding && branding.primaryColor && branding.primaryColor !== '#2B579A') {
    return {
      primary: branding.primaryColor.replace('#', ''),
      secondary: branding.secondaryColor.replace('#', ''),
      accent: branding.accentColor.replace('#', ''),
    };
  }
  return TEMPLATES[template].colors;
}

// Template configurations
const TEMPLATES: Record<TemplateType, { colors: { primary: string; secondary: string; accent: string }; icon: string; title: string }> = {
  'professional-report': { colors: { primary: '2B579A', secondary: '1F4E79', accent: '5B9BD5' }, icon: '📋', title: 'Professional Report' },
  'meeting-summary': { colors: { primary: '217346', secondary: '185C37', accent: '33C481' }, icon: '📅', title: 'Meeting Summary' },
  'project-status': { colors: { primary: 'B7472A', secondary: '8C3620', accent: 'E35B3F' }, icon: '📊', title: 'Project Status' },
  'data-analysis': { colors: { primary: '4472C4', secondary: '305496', accent: '8FAADC' }, icon: '📈', title: 'Data Analysis' },
  'sales-pitch': { colors: { primary: 'FFC000', secondary: 'CC9900', accent: 'FFD966' }, icon: '💼', title: 'Sales Pitch' },
  'email-summary': { colors: { primary: '7030A0', secondary: '5B2680', accent: '9B59B6' }, icon: '✉️', title: 'Email Summary' },
  'action-items': { colors: { primary: '00B050', secondary: '008A3E', accent: '92D050' }, icon: '✅', title: 'Action Items' },
  'custom': { colors: { primary: '39FF14', secondary: '32CD32', accent: 'FFB000' }, icon: '🤖', title: 'AI Generated' },
};

/**
 * Parse markdown content into structured data
 */
function parseContent(content: string) {
  const lines = content.split('\n');
  const paragraphs: string[] = [];
  const tables: string[][] = [];
  const numbers: { label: string; value: number }[] = [];
  let currentTable: string[] = [];
  let inTable = false;
  
  for (const line of lines) {
    // Extract numbers for potential charts
    const numberMatch = line.match(/^[•\-\*]?\s*(.+?):\s*([\d,]+\.?\d*)/);
    if (numberMatch) {
      numbers.push({ label: numberMatch[1].trim(), value: parseFloat(numberMatch[2].replace(/,/g, '')) });
    }
    
    if (line.includes('|') && line.trim().startsWith('|')) {
      inTable = true;
      if (!line.includes('---')) {
        currentTable.push(line.split('|').filter(c => c.trim()).map(c => c.trim()).join('\t'));
      }
    } else {
      if (inTable && currentTable.length > 0) {
        tables.push(currentTable);
        currentTable = [];
        inTable = false;
      }
      if (line.trim()) paragraphs.push(line.trim());
    }
  }
  if (currentTable.length > 0) tables.push(currentTable);
  return { paragraphs, tables, numbers };
}

/**
 * Auto-detect best template based on content
 */
function detectTemplate(content: string): TemplateType {
  const lower = content.toLowerCase();
  if (lower.includes('meeting') || lower.includes('attendees') || lower.includes('agenda')) return 'meeting-summary';
  if (lower.includes('action item') || lower.includes('task') || lower.includes('todo')) return 'action-items';
  if (lower.includes('project') || lower.includes('status') || lower.includes('milestone')) return 'project-status';
  if (lower.includes('analysis') || lower.includes('data') || lower.includes('chart')) return 'data-analysis';
  if (lower.includes('email') || lower.includes('inbox') || lower.includes('message')) return 'email-summary';
  if (lower.includes('sales') || lower.includes('pitch') || lower.includes('proposal')) return 'sales-pitch';
  if (lower.includes('report') || lower.includes('executive') || lower.includes('summary')) return 'professional-report';
  return 'custom';
}

/**
 * Generate Word document (.docx) with templates and enhanced formatting
 */
async function generateWord(title: string, content: string, options: DocumentOptions = {}): Promise<GeneratedDocument> {
  const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType, Header, Footer, PageNumber, ImageRun } = await loadDocx();
  
  const template = options.template || detectTemplate(content);
  const branding = options.useCustomBranding !== false ? loadBranding() : undefined;
  const colors = getColors(template, branding);
  const { paragraphs, tables } = parseContent(content);
  
  const children: any[] = [];
  
  // Add logo if user has one
  if (branding?.logoUrl && branding.logoUrl.startsWith('data:image')) {
    try {
      const base64Data = branding.logoUrl.split(',')[1];
      const imageBuffer = Uint8Array.from(atob(base64Data), c => c.charCodeAt(0));
      children.push(
        new Paragraph({
          children: [
            new ImageRun({
              data: imageBuffer,
              transformation: { width: branding.logoWidth, height: branding.logoWidth * 0.5 },
              type: 'png',
            }),
          ],
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
        })
      );
    } catch (e) {
      console.warn('Failed to add logo to Word doc:', e);
    }
  }
  
  // Company name from branding
  if (branding?.companyName) {
    children.push(
      new Paragraph({
        children: [new TextRun({ text: branding.companyName, bold: true, size: 32, color: colors.primary })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 100 },
      })
    );
  }
  
  // Add Table of Contents placeholder if requested
  if (options.includeTableOfContents) {
    children.push(
      new Paragraph({ text: 'Table of Contents', heading: HeadingLevel.HEADING_1, spacing: { after: 200 } }),
      new Paragraph({ text: '(Update field to generate TOC)', spacing: { after: 400 }, style: 'Normal' })
    );
  }
  
  // Title with template styling
  children.push(
    new Paragraph({
      children: [
        new TextRun({ text: TEMPLATES[template].icon + ' ', size: 48 }),
        new TextRun({ text: title, bold: true, size: 48, color: colors.primary }),
      ],
      spacing: { after: 200 },
      alignment: AlignmentType.CENTER,
    }),
    new Paragraph({
      children: [new TextRun({ text: `Generated: ${new Date().toLocaleString()}`, italics: true, size: 20, color: '888888' })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
    })
  );
  
  // Content with enhanced formatting
  let inBulletList = false;
  for (const para of paragraphs) {
    let text = para;
    let heading: any = undefined;
    
    if (para.startsWith('### ')) { heading = HeadingLevel.HEADING_3; text = para.slice(4); }
    else if (para.startsWith('## ')) { heading = HeadingLevel.HEADING_2; text = para.slice(3); }
    else if (para.startsWith('# ')) { heading = HeadingLevel.HEADING_1; text = para.slice(2); }
    
    if (heading) {
      children.push(new Paragraph({ 
        children: [new TextRun({ text: text.replace(/\*\*/g, ''), bold: true, color: colors.secondary })],
        heading, 
        spacing: { before: 300, after: 150 } 
      }));
      inBulletList = false;
    } else if (para.startsWith('• ') || para.startsWith('- ') || para.startsWith('* ') || para.match(/^\d+\./)) {
      const bulletText = para.replace(/^[•\-\*\d.]+\s*/, '');
      children.push(new Paragraph({ 
        children: [new TextRun({ text: bulletText.replace(/\*\*/g, '') })], 
        bullet: { level: 0 }, 
        spacing: { after: 80 } 
      }));
      inBulletList = true;
    } else {
      // Parse bold text
      const parts: any[] = [];
      const boldRegex = /\*\*(.*?)\*\*/g;
      let lastIndex = 0;
      let match;
      while ((match = boldRegex.exec(text)) !== null) {
        if (match.index > lastIndex) {
          parts.push(new TextRun({ text: text.slice(lastIndex, match.index) }));
        }
        parts.push(new TextRun({ text: match[1], bold: true, color: colors.accent }));
        lastIndex = match.index + match[0].length;
      }
      if (lastIndex < text.length) {
        parts.push(new TextRun({ text: text.slice(lastIndex) }));
      }
      children.push(new Paragraph({ children: parts.length > 0 ? parts : [new TextRun({ text })], spacing: { after: 150 } }));
    }
  }
  
  // Tables with styling
  if (tables.length > 0) {
    const tableData = tables[0].map(row => row.split('\t'));
    if (tableData.length > 0) {
      children.push(new Paragraph({ spacing: { before: 200 } }));
      children.push(new Table({
        rows: tableData.map((row, idx) => new TableRow({
          children: row.map(cell => new TableCell({
            children: [new Paragraph({ 
              children: [new TextRun({ text: cell, bold: idx === 0, color: idx === 0 ? 'FFFFFF' : undefined })] 
            })],
            width: { size: Math.floor(100 / row.length), type: WidthType.PERCENTAGE },
            shading: idx === 0 ? { fill: colors.primary } : undefined,
            borders: {
              top: { style: BorderStyle.SINGLE, size: 1, color: colors.primary },
              bottom: { style: BorderStyle.SINGLE, size: 1, color: colors.primary },
              left: { style: BorderStyle.SINGLE, size: 1, color: colors.primary },
              right: { style: BorderStyle.SINGLE, size: 1, color: colors.primary },
            },
          })),
        })),
        width: { size: 100, type: WidthType.PERCENTAGE },
      }));
    }
  }
  
  // Add contact info from branding
  if (branding && (branding.email || branding.phone || branding.website)) {
    children.push(new Paragraph({ spacing: { before: 400 } }));
    children.push(new Paragraph({
      children: [new TextRun({ text: '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', color: colors.primary })],
      alignment: AlignmentType.CENTER,
    }));
    
    const contactLines = [];
    if (branding.contactName) contactLines.push(`${branding.contactName}${branding.title ? ` | ${branding.title}` : ''}`);
    if (branding.email) contactLines.push(`📧 ${branding.email}`);
    if (branding.phone) contactLines.push(`📞 ${branding.phone}`);
    if (branding.website) contactLines.push(`🌐 ${branding.website}`);
    const address = getFormattedAddress(branding);
    if (address) contactLines.push(`📍 ${address}`);
    
    contactLines.forEach(line => {
      children.push(new Paragraph({
        children: [new TextRun({ text: line, size: 20, color: '666666' })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 50 },
      }));
    });
  }
  
  // Footer
  const footerText = branding?.companyName 
    ? `${TEMPLATES[template].icon} Generated by ${branding.companyName} using Outlook AI`
    : `${TEMPLATES[template].icon} Generated by Outlook AI - FreedomForged_AI`;
  
  children.push(new Paragraph({
    children: [new TextRun({ text: footerText, italics: true, size: 18, color: '888888' })],
    spacing: { before: 600 },
    alignment: AlignmentType.CENTER,
  }));
  
  const doc = new Document({
    creator: options.author || 'Outlook AI - FreedomForged_AI',
    title: title,
    description: `${TEMPLATES[template].title} generated by Outlook AI`,
    sections: [{
      properties: {},
      children,
      headers: {
        default: new Header({
          children: [new Paragraph({ 
            children: [new TextRun({ text: title, size: 16, color: '888888' })],
            alignment: AlignmentType.RIGHT,
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            children: [
              new TextRun({ text: 'Page ', size: 16, color: '888888' }),
              new TextRun({ children: [PageNumber.CURRENT], size: 16, color: '888888' }),
              new TextRun({ text: ' | Outlook AI - FreedomForged_AI', size: 16, color: '888888' }),
            ],
            alignment: AlignmentType.CENTER,
          })],
        }),
      },
    }],
  });
  
  const blob = await Packer.toBlob(doc);
  const filename = `${TEMPLATES[template].title.replace(/\s/g, '_')}_${title.replace(/[^a-z0-9]/gi, '_')}.docx`;
  return { blob, filename, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' };
}

/**
 * Generate HTML (printable to PDF) with enhanced styling
 */
async function generatePDF(title: string, content: string, options: DocumentOptions = {}): Promise<GeneratedDocument> {
  const template = options.template || detectTemplate(content);
  const branding = options.useCustomBranding !== false ? loadBranding() : undefined;
  const colors = getColors(template, branding);
  const { paragraphs, tables, numbers } = parseContent(content);
  
  // Generate SVG chart if we have numerical data
  let chartSvg = '';
  if (numbers.length >= 2 && options.includeCharts !== false) {
    const maxValue = Math.max(...numbers.map(n => n.value));
    const barHeight = 30;
    const chartHeight = numbers.length * (barHeight + 10) + 40;
    chartSvg = `
    <div class="chart-container">
      <h3>📊 Data Visualization</h3>
      <svg width="100%" height="${chartHeight}" viewBox="0 0 500 ${chartHeight}">
        ${numbers.map((n, i) => {
          const width = (n.value / maxValue) * 350;
          const y = i * (barHeight + 10) + 30;
          return `
            <text x="0" y="${y + 20}" font-size="12" fill="#666">${n.label.substring(0, 20)}</text>
            <rect x="130" y="${y}" width="${width}" height="${barHeight}" fill="#${colors.primary}" rx="4"/>
            <text x="${135 + width}" y="${y + 20}" font-size="12" fill="#333">${n.value.toLocaleString()}</text>
          `;
        }).join('')}
      </svg>
    </div>`;
  }
  
  // Build contact info section
  const contactHtml = branding && (branding.email || branding.phone) ? `
    <div class="contact-info">
      ${branding.contactName ? `<p><strong>${branding.contactName}</strong>${branding.title ? ` | ${branding.title}` : ''}</p>` : ''}
      ${branding.email ? `<p>📧 ${branding.email}</p>` : ''}
      ${branding.phone ? `<p>📞 ${branding.phone}</p>` : ''}
      ${branding.website ? `<p>🌐 <a href="${branding.website}">${branding.website}</a></p>` : ''}
      ${getFormattedAddress(branding) ? `<p>📍 ${getFormattedAddress(branding)}</p>` : ''}
    </div>
  ` : '';
  
  let html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>${title}</title>
<style>
  @page { margin: 2cm; }
  body{font-family:'Segoe UI',Arial,sans-serif;max-width:800px;margin:40px auto;padding:20px;line-height:1.7;background:#fff;color:#333}
  .logo{text-align:center;margin-bottom:20px}
  .logo img{max-width:200px;height:auto}
  .company-name{text-align:center;font-size:24px;font-weight:bold;color:#${colors.primary};margin-bottom:10px}
  .header{text-align:center;border-bottom:3px solid #${colors.primary};padding-bottom:20px;margin-bottom:30px}
  .header h1{color:#${colors.primary};margin:0 0 10px 0}
  .header .meta{color:#888;font-size:12px}
  h2{color:#${colors.secondary};margin-top:30px;border-left:4px solid #${colors.primary};padding-left:15px}
  h3{color:#${colors.accent};margin-top:20px}
  table{width:100%;border-collapse:collapse;margin:20px 0;box-shadow:0 2px 8px rgba(0,0,0,0.1)}
  th{background:#${colors.primary};color:#fff;padding:12px;text-align:left;font-weight:600}
  td{border:1px solid #e0e0e0;padding:12px;text-align:left}
  tr:nth-child(even){background:#f9f9f9}
  tr:hover{background:#f0f0f0}
  ul{margin:15px 0;padding-left:25px}
  li{margin:8px 0}
  .highlight{background:linear-gradient(120deg,rgba(${parseInt(colors.accent.slice(0,2),16)},${parseInt(colors.accent.slice(2,4),16)},${parseInt(colors.accent.slice(4,6),16)},0.2) 0%,transparent 100%);padding:2px 6px;border-radius:4px}
  .chart-container{background:#f8f9fa;padding:20px;border-radius:8px;margin:20px 0}
  .chart-container h3{margin:0 0 15px 0;color:#${colors.primary}}
  .contact-info{margin-top:30px;padding:20px;background:#f8f9fa;border-radius:8px;border-left:4px solid #${colors.primary}}
  .contact-info p{margin:5px 0;font-size:14px}
  .footer{margin-top:40px;padding-top:20px;border-top:2px solid #${colors.primary};color:#888;font-size:12px;text-align:center}
  .badge{display:inline-block;background:#${colors.primary};color:#fff;padding:4px 12px;border-radius:20px;font-size:11px;margin-right:8px}
  @media print{body{margin:0;background:#fff}.chart-container{break-inside:avoid}}
</style></head><body>
<div class="header">
  ${branding?.logoUrl ? `<div class="logo"><img src="${branding.logoUrl}" alt="Logo" style="max-width:${branding.logoWidth}px"></div>` : ''}
  ${branding?.companyName ? `<div class="company-name">${branding.companyName}</div>` : ''}
  <h1>${TEMPLATES[template].icon} ${title}</h1>
  <div class="meta">
    <span class="badge">${TEMPLATES[template].title}</span>
    Generated: ${new Date().toLocaleString()}
  </div>
</div>`;
  
  let inList = false;
  for (const para of paragraphs) {
    const isList = para.startsWith('• ') || para.startsWith('- ') || para.startsWith('* ') || para.match(/^\d+\./);
    if (isList && !inList) { html += '<ul>'; inList = true; }
    if (!isList && inList) { html += '</ul>'; inList = false; }
    
    if (para.startsWith('### ')) html += `<h3>${para.slice(4)}</h3>`;
    else if (para.startsWith('## ')) html += `<h2>${para.slice(3)}</h2>`;
    else if (para.startsWith('# ')) html += `<h1>${para.slice(2)}</h1>`;
    else if (isList) html += `<li>${para.replace(/^[•\-\*\d.]+\s*/, '').replace(/\*\*(.*?)\*\*/g, '<strong class="highlight">$1</strong>')}</li>`;
    else html += `<p>${para.replace(/\*\*(.*?)\*\*/g, '<strong class="highlight">$1</strong>')}</p>`;
  }
  if (inList) html += '</ul>';
  
  // Add chart
  if (chartSvg) html += chartSvg;
  
  // Add tables
  if (tables.length > 0) {
    const tableData = tables[0].map(row => row.split('\t'));
    if (tableData.length > 0) {
      html += '<table><thead><tr>';
      tableData[0].forEach(cell => html += `<th>${cell}</th>`);
      html += '</tr></thead><tbody>';
      tableData.slice(1).forEach(row => {
        html += '<tr>'; row.forEach(cell => html += `<td>${cell}</td>`); html += '</tr>';
      });
      html += '</tbody></table>';
    }
  }
  
  // Add contact info
  if (contactHtml) html += contactHtml;
  
  const footerText = branding?.companyName 
    ? `${TEMPLATES[template].icon} Generated by <strong>${branding.companyName}</strong> using Outlook AI`
    : `${TEMPLATES[template].icon} Generated by <strong>Outlook AI - FreedomForged_AI</strong>`;
  
  html += `<div class="footer">
    ${footerText}<br>
    <small>To save as PDF: Press Ctrl+P (or Cmd+P) → Destination: "Save as PDF"</small>
  </div></body></html>`;
  
  const filename = `${TEMPLATES[template].title.replace(/\s/g, '_')}_${title.replace(/[^a-z0-9]/gi, '_')}.html`;
  return { blob: new Blob([html], { type: 'text/html' }), filename, mimeType: 'text/html' };
}
async function generateExcel(title: string, content: string, options: DocumentOptions = {}): Promise<GeneratedDocument> {
  const ExcelJS = await loadExcel();
  const template = options.template || detectTemplate(content);
  const colors = TEMPLATES[template].colors;
  
  const workbook = new ExcelJS.Workbook();
  workbook.creator = options.author || 'Outlook AI - FreedomForged_AI';
  workbook.created = new Date();
  
  const { tables, paragraphs, numbers } = parseContent(content);
  const tableData = tables.length > 0 ? tables[0].map(row => row.split('\t')) : null;
  
  // Main data sheet
  const dataSheet = workbook.addWorksheet(title.substring(0, 31));
  
  // Title row with styling
  const titleRow = dataSheet.addRow([`${TEMPLATES[template].icon} ${title}`]);
  titleRow.font = { bold: true, size: 18, color: { argb: 'FF' + colors.primary } };
  titleRow.height = 30;
  dataSheet.mergeCells('A1:F1');
  
  dataSheet.addRow([`Generated: ${new Date().toLocaleString()}`]).font = { italic: true, color: { argb: 'FF888888' } };
  dataSheet.addRow([]);
  
  if (tableData && tableData.length > 0) {
    // Add table data with formatting
    tableData.forEach((row, idx) => {
      const excelRow = dataSheet.addRow(row);
      if (idx === 0) {
        excelRow.eachCell(cell => {
          cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + colors.primary } };
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FF' + colors.primary } },
            bottom: { style: 'thin', color: { argb: 'FF' + colors.primary } },
            left: { style: 'thin', color: { argb: 'FF' + colors.primary } },
            right: { style: 'thin', color: { argb: 'FF' + colors.primary } },
          };
        });
      } else {
        excelRow.eachCell(cell => {
          cell.border = {
            top: { style: 'thin', color: { argb: 'FFE0E0E0' } },
            bottom: { style: 'thin', color: { argb: 'FFE0E0E0' } },
            left: { style: 'thin', color: { argb: 'FFE0E0E0' } },
            right: { style: 'thin', color: { argb: 'FFE0E0E0' } },
          };
          if (idx % 2 === 0) {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5F5F5' } };
          }
        });
      }
    });
    
    // Auto-width columns
    dataSheet.columns.forEach(col => {
      col.width = 18;
    });
  } else {
    // No table - add content as rows
    paragraphs.forEach(line => {
      const row = dataSheet.addRow([line]);
      if (line.startsWith('#')) {
        row.font = { bold: true, size: 14, color: { argb: 'FF' + colors.primary } };
      } else if (line.startsWith('• ') || line.startsWith('- ')) {
        row.getCell(1).alignment = { indent: 2 };
      }
    });
    dataSheet.getColumn(1).width = 100;
  }
  
  // Add chart data sheet if we have numbers
  if (numbers.length >= 2 && options.includeCharts !== false) {
    const chartSheet = workbook.addWorksheet('Chart Data');
    chartSheet.addRow(['📊 Data for Charts']);
    chartSheet.addRow([]);
    chartSheet.addRow(['Label', 'Value']);
    numbers.forEach(n => chartSheet.addRow([n.label, n.value]));
    
    // Style the chart data
    chartSheet.getRow(3).eachCell(cell => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + colors.primary } };
    });
    chartSheet.getColumn(1).width = 30;
    chartSheet.getColumn(2).width = 15;
  }
  
  // Footer
  dataSheet.addRow([]);
  const footerRow = dataSheet.addRow([`${TEMPLATES[template].icon} Generated by Outlook AI - FreedomForged_AI`]);
  footerRow.font = { italic: true, color: { argb: 'FF888888' }, size: 10 };
  
  const buffer = await workbook.xlsx.writeBuffer();
  const filename = `${TEMPLATES[template].title.replace(/\s/g, '_')}_${title.replace(/[^a-z0-9]/gi, '_')}.xlsx`;
  return { blob: new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), filename, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' };
}

/**
 * Generate PowerPoint-like HTML presentation with enhanced animations
 */
async function generatePowerPoint(title: string, content: string, options: DocumentOptions = {}): Promise<GeneratedDocument> {
  const template = options.template || detectTemplate(content);
  const colors = TEMPLATES[template].colors;
  const { paragraphs, tables, numbers } = parseContent(content);
  const slides: { title: string; bullets: string[]; type: 'content' | 'data' }[] = [];
  let currentSlide = { title: title, bullets: [] as string[], type: 'content' as const };
  
  for (const para of paragraphs) {
    if (para.startsWith('## ') || para.startsWith('# ')) {
      if (currentSlide.bullets.length > 0 || currentSlide.title !== title) slides.push(currentSlide);
      currentSlide = { title: para.replace(/^#+ /, ''), bullets: [], type: 'content' };
    } else if (!para.startsWith('### ')) {
      currentSlide.bullets.push(para.replace(/\*\*/g, ''));
      if (currentSlide.bullets.length >= 5) {
        slides.push(currentSlide);
        currentSlide = { title: currentSlide.title + ' (cont.)', bullets: [], type: 'content' };
      }
    }
  }
  if (currentSlide.bullets.length > 0) slides.push(currentSlide);
  
  // Chart SVG for data slides
  let chartSvg = '';
  if (numbers.length >= 2) {
    const maxValue = Math.max(...numbers.map(n => n.value));
    chartSvg = numbers.map((n, i) => {
      const width = (n.value / maxValue) * 80;
      return `<div class="bar-item"><span class="bar-label">${n.label}</span><div class="bar" style="width:${width}%"><span class="bar-value">${n.value.toLocaleString()}</span></div></div>`;
    }).join('');
  }
  
  let html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>${title}</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  html,body{height:100%;overflow:hidden}
  body{font-family:'Segoe UI',Arial;background:#1a1a1a;color:#e0e0e0}
  .slide{width:100vw;height:100vh;padding:60px 80px;display:none;flex-direction:column;page-break-after:always;position:relative;overflow:hidden}
  .slide.active{display:flex}
  .slide::before{content:'';position:absolute;top:0;right:0;width:300px;height:300px;background:radial-gradient(circle,rgba(${parseInt(colors.primary.slice(0,2),16)},${parseInt(colors.primary.slice(2,4),16)},${parseInt(colors.primary.slice(4,6),16)},0.1) 0%,transparent 70%);pointer-events:none}
  .slide h1{color:#${colors.primary};font-size:56px;margin-bottom:40px;text-shadow:0 0 30px rgba(${parseInt(colors.primary.slice(0,2),16)},${parseInt(colors.primary.slice(2,4),16)},${parseInt(colors.primary.slice(4,6),16)},0.5);animation:fadeInUp 0.5s ease}
  .slide h2{color:#${colors.secondary};font-size:42px;margin-bottom:35px;border-left:5px solid #${colors.primary};padding-left:20px;animation:fadeInUp 0.5s ease}
  .slide ul{font-size:26px;line-height:2.2;padding-left:40px;flex:1}
  .slide li{margin:15px 0;animation:fadeInLeft 0.5s ease;animation-fill-mode:both}
  .slide li:nth-child(1){animation-delay:0.1s}
  .slide li:nth-child(2){animation-delay:0.2s}
  .slide li:nth-child(3){animation-delay:0.3s}
  .slide li:nth-child(4){animation-delay:0.4s}
  .slide li:nth-child(5){animation-delay:0.5s}
  .title-slide{text-align:center;justify-content:center;background:linear-gradient(135deg,#1a1a1a 0%,#2a2a3a 100%)}
  .title-slide h1{font-size:72px;margin-bottom:30px}
  .title-slide .subtitle{color:#${colors.accent};font-size:28px;margin-top:20px}
  .title-slide .meta{color:#888;font-size:18px;margin-top:40px}
  .end-slide{text-align:center;justify-content:center;background:linear-gradient(135deg,#2a2a3a 0%,#1a1a1a 100%)}
  .end-slide h1{color:#${colors.primary};font-size:64px}
  .footer{position:absolute;bottom:30px;left:80px;right:80px;display:flex;justify-content:space-between;color:#666;font-size:14px}
  .slide-number{background:#${colors.primary};color:#000;padding:5px 15px;border-radius:20px;font-weight:bold}
  .nav-hint{position:fixed;bottom:20px;right:20px;background:rgba(0,0,0,0.8);padding:10px 20px;border-radius:8px;font-size:12px;color:#888}
  .bar-chart{background:rgba(0,0,0,0.3);padding:30px;border-radius:12px;margin-top:20px}
  .bar-item{display:flex;align-items:center;margin:15px 0}
  .bar-label{width:150px;font-size:16px;color:#aaa}
  .bar{background:linear-gradient(90deg,#${colors.primary},#${colors.accent});height:35px;border-radius:4px;display:flex;align-items:center;justify-content:flex-end;padding-right:10px;transition:width 0.5s ease}
  .bar-value{color:#000;font-weight:bold;font-size:14px}
  table{width:100%;border-collapse:collapse;margin-top:20px;animation:fadeIn 0.5s ease}
  th{background:#${colors.primary};color:#000;padding:15px;text-align:left;font-size:18px}
  td{padding:15px;border-bottom:1px solid #333;font-size:16px}
  tr:hover{background:rgba(${parseInt(colors.primary.slice(0,2),16)},${parseInt(colors.primary.slice(2,4),16)},${parseInt(colors.primary.slice(4,6),16)},0.1)}
  @keyframes fadeInUp{from{opacity:0;transform:translateY(30px)}to{opacity:1;transform:translateY(0)}}
  @keyframes fadeInLeft{from{opacity:0;transform:translateX(-30px)}to{opacity:1;transform:translateX(0)}}
  @keyframes fadeIn{from{opacity:0}to{opacity:1}}
  @media print{.slide{display:flex!important;height:100vh;page-break-after:always}.nav-hint{display:none}}
</style></head><body>

<div class="slide title-slide active" data-slide="0">
  <h1>${TEMPLATES[template].icon} ${title}</h1>
  <div class="subtitle">${TEMPLATES[template].title}</div>
  <div class="meta">Generated by Outlook AI - FreedomForged_AI<br>${new Date().toLocaleString()}</div>
</div>`;
  
  slides.forEach((slide, idx) => {
    html += `<div class="slide" data-slide="${idx + 1}">
      <h2>${slide.title}</h2>
      <ul>${slide.bullets.map(b => `<li>${b}</li>`).join('')}</ul>
      <div class="footer">
        <span>Outlook AI - FreedomForged_AI</span>
        <span class="slide-number">${idx + 2} / ${slides.length + (numbers.length >= 2 ? 3 : 2) + (tables.length > 0 ? 1 : 0)}</span>
      </div>
    </div>`;
  });
  
  // Data slide with chart
  if (numbers.length >= 2) {
    html += `<div class="slide" data-slide="${slides.length + 1}">
      <h2>📊 Key Metrics</h2>
      <div class="bar-chart">${chartSvg}</div>
      <div class="footer">
        <span>Outlook AI - FreedomForged_AI</span>
        <span class="slide-number">${slides.length + 2} / ${slides.length + 3 + (tables.length > 0 ? 1 : 0)}</span>
      </div>
    </div>`;
  }
  
  // Table slide
  if (tables.length > 0) {
    const tableData = tables[0].map(row => row.split('\t'));
    html += `<div class="slide" data-slide="${slides.length + (numbers.length >= 2 ? 2 : 1)}">
      <h2>📋 Data Table</h2>
      <table>
        <thead><tr>${tableData[0].map(c => `<th>${c}</th>`).join('')}</tr></thead>
        <tbody>${tableData.slice(1).map(row => `<tr>${row.map(c => `<td>${c}</td>`).join('')}</tr>`).join('')}</tbody>
      </table>
      <div class="footer">
        <span>Outlook AI - FreedomForged_AI</span>
        <span class="slide-number">${slides.length + (numbers.length >= 2 ? 3 : 2)} / ${slides.length + (numbers.length >= 2 ? 3 : 2) + (tables.length > 0 ? 1 : 0)}</span>
      </div>
    </div>`;
  }
  
  html += `<div class="slide end-slide" data-slide="end">
    <h1>Thank You! 🙏</h1>
    <div class="subtitle" style="margin-top:30px">Questions?</div>
    <div class="meta" style="margin-top:40px">Powered by <strong>Outlook AI</strong> - FreedomForged_AI</div>
  </div>

  <div class="nav-hint">← → Arrow keys to navigate | Press F11 for fullscreen</div>

  <script>
    let current=0;
    const slides=document.querySelectorAll('.slide');
    const total=slides.length;
    
    function showSlide(n){
      slides.forEach(s=>s.classList.remove('active'));
      current=((n%total)+total)%total;
      slides[current].classList.add('active');
    }
    
    document.addEventListener('keydown',e=>{
      if(e.key==='ArrowRight'||e.key===' ')showSlide(current+1);
      if(e.key==='ArrowLeft')showSlide(current-1);
      if(e.key==='Home')showSlide(0);
      if(e.key==='End')showSlide(total-1);
    });
    
    document.addEventListener('click',e=>{
      if(e.clientX>window.innerWidth/2)showSlide(current+1);
      else showSlide(current-1);
    });
  </script>
</body></html>`;
  
  const filename = `${TEMPLATES[template].title.replace(/\s/g, '_')}_${title.replace(/[^a-z0-9]/gi, '_')}_presentation.html`;
  return { blob: new Blob([html], { type: 'text/html' }), filename, mimeType: 'text/html' };
}

/**
 * Download document
 */
export function downloadDocument(doc: GeneratedDocument): void {
  const url = URL.createObjectURL(doc.blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = doc.filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

/**
 * Open document in new window (for HTML/PDF preview)
 */
export function openDocument(doc: GeneratedDocument): void {
  const url = URL.createObjectURL(doc.blob);
  window.open(url, '_blank');
}

/**
 * Get available templates
 */
export function getTemplates() {
  return Object.entries(TEMPLATES).map(([key, value]) => ({
    id: key as TemplateType,
    name: value.title,
    icon: value.icon,
  }));
}

/**
 * Document service with lazy loading
 */
export const documentService = {
  templates: getTemplates(),
  
  async createWord(title: string, content: string, options?: DocumentOptions) {
    const doc = await generateWord(title, content, options);
    downloadDocument(doc);
    return doc;
  },
  
  async createPDF(title: string, content: string, options?: DocumentOptions) {
    const doc = await generatePDF(title, content, options);
    openDocument(doc); // Open in new tab for print-to-PDF
    return doc;
  },
  
  async createExcel(title: string, content: string, options?: DocumentOptions) {
    const doc = await generateExcel(title, content, options);
    downloadDocument(doc);
    return doc;
  },
  
  async createPowerPoint(title: string, content: string, options?: DocumentOptions) {
    const doc = await generatePowerPoint(title, content, options);
    openDocument(doc); // Open in fullscreen for presentation
    return doc;
  },
  
  // Utility to detect best template
  detectTemplate,
  
  // Pre-load libraries (call on app init for faster first export)
  async preload() {
    await Promise.all([loadDocx(), loadExcel()]);
  },
};

export default documentService;
