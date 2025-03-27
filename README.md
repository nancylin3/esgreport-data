# esgreport-data
與永續報告書資料庫相關
// 永續報告書PDF解析處理程序 (Node.js)

const fs = require('fs');
const path = require('path');
const { PDFDocument } = require('pdf-lib');
const pdf = require('pdf-parse');
const natural = require('natural');
const mongoose = require('mongoose');
const { Worker } = require('worker_threads');
const { createWorker } = require('tesseract.js');
const { Configuration, OpenAIApi } = require('openai');

// 載入相關模型
const Report = require('../models/Report');
const Chapter = require('../models/Chapter');
const Indicator = require('../models/Indicator');
const Goal = require('../models/Goal');

// NLP 相關工具
const tokenizer = new natural.WordTokenizer();
const TfIdf = natural.TfIdf;
const tfidf = new TfIdf();

// OpenAI 配置 (用於進階內容理解)
const configuration = new Configuration({
  apiKey: process.env.OPENAI_API_KEY,
});
const openai = new OpenAIApi(configuration);

// 主要處理函數
async function processReport(reportId) {
  console.log(`開始處理報告 ID: ${reportId}`);
  
  try {
    // 1. 獲取報告信息
    const report = await Report.findById(reportId);
    if (!report) {
      throw new Error(`找不到報告 ID: ${reportId}`);
    }
    
    // 更新處理狀態
    report.processingStatus = 'processing';
    await report.save();
    
    // 2. 讀取PDF文件
    const filePath = path.join(process.cwd(), report.fileUrl);
    const dataBuffer = fs.readFileSync(filePath);
    
    // 3. 解析PDF基本信息
    const pdfData = await pdf(dataBuffer);
    console.log(`PDF總頁數: ${pdfData.numpages}`);
    report.totalPages = pdfData.numpages;
    await report.save();
    
    // 4. 獲取目錄結構
    const tableOfContents = await extractTableOfContents(dataBuffer);
    console.log(`識別到的目錄項: ${tableOfContents.length}`);
    
    // 5. 按章節拆分內容
    const chapters = await splitByChapters(dataBuffer, tableOfContents);
    
    // 6. 處理每個章節
    for (const chapter of chapters) {
      await processChapter(report._id, chapter);
    }
    
    // 7. 提取關鍵績效指標
    await extractIndicators(report._id, dataBuffer);
    
    // 8. 提取永續目標
    await extractSustainabilityGoals(report._id, dataBuffer);
    
    // 9. 更新報告處理狀態為完成
    report.processingStatus = 'completed';
    await report.save();
    
    console.log(`報告 ID: ${reportId} 處理完成`);
    return true;
  } catch (error) {
    console.error(`處理報告時出錯 ID: ${reportId}`, error);
    
    // 更新報告處理狀態為失敗
    const report = await Report.findById(reportId);
    if (report) {
      report.processingStatus = 'failed';
      await report.save();
    }
    
    throw error;
  }
}

// 提取目錄結構
async function extractTableOfContents(pdfBuffer) {
  try {
    // 使用 pdf-parse 獲取文本
    const data = await pdf(pdfBuffer);
    const text = data.text;
    
    // 尋找目錄部分 (假設目錄在前20頁)
    let tocText = '';
    const pdfDoc = await PDFDocument.load(pdfBuffer);
    
    for (let i = 0; i < Math.min(20, pdfDoc.getPageCount()); i++) {
      const page = await extractPageText(pdfBuffer, i);
      
      // 檢查頁面是否包含目錄相關標題
      if (page.includes('目錄') || page.includes('Contents') || page.includes('Table of Contents')) {
        tocText = page;
        break;
      }
    }
    
    if (!tocText) {
      console.warn('無法找到目錄部分，嘗試使用啟發式方法識別章節');
      return extractChaptersHeuristically(text);
    }
    
    // 解析目錄文本
    const tocLines = tocText.split('\n').filter(line => line.trim());
    const chapters = [];
    
    // 識別章節模式 (例如: "1. 章節名稱...... 10" 或 "Chapter 1: 章節名稱...... 10")
    const chapterPattern = /^(?:Chapter\s+)?(\d+(?:\.\d+)*)[\.:\s]+([^\d]+)(?:\.+|\s+)(\d+)$/i;
    
    for (const line of tocLines) {
      const match = line.match(chapterPattern);
      if (match) {
        chapters.push({
          number: match[1],
          title: match[2].trim(),
          page: parseInt(match[3])
        });
      }
    }
    
    // 如果識別到的章節太少，嘗試使用啟發式方法
    if (chapters.length < 3) {
      console.warn('目錄識別章節過少，使用啟發式方法');
      return extractChaptersHeuristically(text);
    }
    
    return chapters;
  } catch (error) {
    console.error('提取目錄時出錯:', error);
    return [];
  }
}

// 啟發式章節識別（當目錄提取失敗時的備選方案）
function extractChaptersHeuristically(text) {
  const chapters = [];
  const lines = text.split('\n');
  
  // 章節標題模式
  const chapterPatterns = [
    /^(\d+(?:\.\d+)*)\s+(.+)$/,                // 數字編號 + 標題 (1. 標題)
    /^(Chapter|Section)\s+(\d+)[:\s]+(.+)$/i,  // Chapter/Section + 數字 + 標題
    /^([A-Z][A-Z\s]+)$/                        // 全大寫標題
  ];
  
  let currentPageEstimate = 1;
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    
    // 頁碼估計 (根據文本行數粗略估計)
    if (i % 50 === 0) {
      currentPageEstimate++;
    }
    
    // 檢查是否符合章節標題模式
    for (const pattern of chapterPatterns) {
      const match = line.match(pattern);
      if (match && line.length < 100) {  // 章節標題通常不會太長
        if (pattern.toString().includes('Chapter|Section')) {
          chapters.push({
            number: match[2],
            title: match[3],
            page: currentPageEstimate
          });
        } else if (pattern.toString().includes('A-Z')) {
          chapters.push({
            number: chapters.length + 1,
            title: match[1],
            page: currentPageEstimate
          });
        } else {
          chapters.push({
            number: match[1],
            title: match[2],
            page: currentPageEstimate
          });
        }
      }
    }
  }
  
  return chapters;
}

// 提取特定頁面的文本
async function extractPageText(pdfBuffer, pageNum) {
  try {
    const options = {
      pagerender: render_page,
      max: pageNum + 1,
      min: pageNum
    };
    
    function render_page(pageData) {
      return pageData.getTextContent()
        .then(function(textContent) {
          let lastY, text = '';
          for (let item of textContent.items) {
            if (lastY == item.transform[5] || !lastY) {
              text += item.str;
            } else {
              text += '\n' + item.str;
            }
            lastY = item.transform[5];
          }
          return text;
        });
    }
    
    const data = await pdf(pdfBuffer, options);
    return data.text;
  } catch (error) {
    console.error(`提取頁面 ${pageNum} 文本時出錯:`, error);
    return '';
  }
}

// 按章節拆分內容
async function splitByChapters(pdfBuffer, tableOfContents) {
  const chaptersData = [];
  
  // 如果沒有識別到章節，返回空數組
  if (!tableOfContents || tableOfContents.length === 0) {
    console.warn('沒有識別到章節，無法拆分');
    return chaptersData;
  }
  
  try {
    // 按頁碼排序章節
    tableOfContents.sort((a, b) => a.page - b.page);
    
    // 逐章節提取內容
    for (let i = 0; i < tableOfContents.length; i++) {
      const chapter = tableOfContents[i];
      const nextChapter = tableOfContents[i + 1];
      
      const startPage = chapter.page - 1; // PDF頁碼從0開始
      const endPage = nextChapter ? nextChapter.page - 2 : null; // 減2是為了避免頁眉頁腳干擾
      
      // 提取章節文本
      let chapterText = '';
      
      if (endPage === null) {
        // 如果是最後一章，讀取到文檔結束
        const data = await pdf(pdfBuffer, { 
          min: startPage,
          max: 1000 // 一個足夠大的數字以確保讀到文檔結束
        });
        chapterText = data.text;
      } else {
        // 否則讀取到下一章開始
        const data = await pdf(pdfBuffer, { 
          min: startPage,
          max: endPage
        });
        chapterText = data.text;
      }
      
      // 儲存章節數據
      chaptersData.push({
        number: chapter.number,
        title: chapter.title,
        startPage: startPage + 1,
        endPage: endPage ? endPage + 1 : null,
        content: chapterText
      });
    }
    
    return chaptersData;
  } catch (error) {
    console.error('按章節拆分內容時出錯:', error);
    return [];
  }
}

// 處理單個章節
async function processChapter(reportId, chapterData) {
  try {
    // 1. 確定章節類型 (ESG分類)
    const chapterType = determineChapterType(chapterData.title, chapterData.content);
    
    // 2. 生成章節摘要
    const contentSummary = await generateChapterSummary(chapterData.content);
    
    // 3. 保存章節數據到數據庫
    const chapter = new Chapter({
      reportId,
      title: chapterData.title,
      chapterNumber: chapterData.number,
      chapterType,
      startPage: chapterData.startPage,
      endPage: chapterData.endPage,
      content: chapterData.content,
      contentSummary
    });
    
    await chapter.save();
    console.log(`保存章節: ${chapterData.title}`);
    
    return chapter._id;
  } catch (error) {
    console.error(`處理章節時出錯: ${chapterData.title}`, error);
    throw error;
  }
}

// 確定章節類型 (環境、社會、治理)
function determineChapterType(title, content) {
  // 常見環境主題關鍵詞
  const environmentalKeywords = [
    '環境', '氣候', '碳', '排放', '能源', '水資源', '廢棄物', '污染', '生物多樣性',
    'environmental', 'climate', 'carbon', 'emission', 'energy', 'water', 'waste', 'pollution', 'biodiversity'
  ];
  
  // 常見社會主題關鍵詞
  const socialKeywords = [
    '社會', '員工', '勞工', '人權', '多元', '包容', '社區', '健康', '安全', '培訓',
    'social', 'employee', 'labor', 'human rights', 'diversity', 'inclusion', 'community', 'health', 'safety', 'training'
  ];
  
  // 常見治理主題關鍵詞
  const governanceKeywords = [
    '治理', '董事會', '風險', '合規', '道德', '反貪腐', '透明', '資訊揭露',
    'governance', 'board', 'risk', 'compliance', 'ethics', 'anti-corruption', 'transparency', 'disclosure'
  ];
  
  // 計算關鍵詞匹配次數
  let eScore = 0, sScore = 0, gScore = 0;
  
  // 檢查標題
  const normalizedTitle = title.toLowerCase();
  
  environmentalKeywords.forEach(keyword => {
    if (normalizedTitle.includes(keyword.toLowerCase())) eScore += 3;
  });
  
  socialKeywords.forEach(keyword => {
    if (normalizedTitle.includes(keyword.toLowerCase())) sScore += 3;
  });
  
  governanceKeywords.forEach(keyword => {
    if (normalizedTitle.includes(keyword.toLowerCase())) gScore += 3;
  });
  
  // 檢查內容
  const normalizedContent = content.toLowerCase();
  
  environmentalKeywords.forEach(keyword => {
    const count = (normalizedContent.match(new RegExp(keyword.toLowerCase(), 'g')) || []).length;
    eScore += count;
  });
  
  socialKeywords.forEach(keyword => {
    const count = (normalizedContent.match(new RegExp(keyword.toLowerCase(), 'g')) || []).length;
    sScore += count;
  });
  
  governanceKeywords.forEach(keyword => {
    const count = (normalizedContent.match(new RegExp(keyword.toLowerCase(), 'g')) || []).length;
    gScore += count;
  });
  
  // 根據得分判斷類型
  if (eScore > sScore && eScore > gScore) return 'E';
  if (sScore > eScore && sScore > gScore) return 'S';
  if (gScore > eScore && gScore > sScore) return 'G';
  
  // 如果得分相同或無法確定，檢查標題是否包含明確的關鍵詞
  if (normalizedTitle.includes('環境') || normalizedTitle.includes('environmental')) return 'E';
  if (normalizedTitle.includes('社會') || normalizedTitle.includes('social')) return 'S';
  if (normalizedTitle.includes('治理') || normalizedTitle.includes('governance')) return 'G';
  
  // 預設為綜合類型
  return 'General';
}

// 生成章節摘要
async function generateChapterSummary(content) {
  // 如果內容過長，截取前2000字符進行處理
  const truncatedContent = content.length > 2000 ? content.substring(0, 2000) : content;
  
  try {
    // 使用 OpenAI API 生成摘要
    const response = await openai.createCompletion({
      model: "text-davinci-003",
      prompt: `請用繁體中文總結以下永續報告書章節的主要內容（100字以內）:\n\n${truncatedContent}`,
      max_tokens: 150,
      temperature: 0.3,
    });
    
    return response.data.choices[0].text.trim();
  } catch (error) {
    console.error('使用 OpenAI 生成摘要時出錯:', error);
    
    // 備用方案：提取前100個字符作為摘要
    return truncatedContent.substring(0, 100) + '...';
  }
}

// 提取關鍵績效指標 (KPIs)
async function extractIndicators(reportId, pdfBuffer) {
  try {
    // 獲取報告全文
    const data = await pdf(pdfBuffer);
    const fullText = data.text;
    
    // 載入已保存的章節資訊
    const chapters = await Chapter.find({ reportId });
    
    // 常見ESG指標模式
    const indicatorPatterns = [
      // GRI 標準模式
      {
        pattern: /GRI\s+(\d+(?:-\d+)?)\s*[:：]?\s*([^\n\d]+)(?:\n|$)/g,
        standard: 'GRI'
      },
      // SASB 標準模式
      {
        pattern: /SASB\s+([A-Z]+(?:-\d+)?)\s*[:：]?\s*([^\n\d]+)(?:\n|$)/g,
        standard: 'SASB'
      },
      // 數值模式 (例如: 範疇一排放量: 2,500,000 公噸CO2e)
      {
        pattern: /([^：:：\n]{2,30})[:：:]\s*([\d,\.]+)\s*([^\d\s\n][^\n]{0,20})/g,
        standard: 'Custom'
      }
    ];
    
    const extractedIndicators = [];
    
    // 從每個章節中提取指標
    for (const chapter of chapters) {
      const pageRange = chapter.content;
      
      // 應用各種模式
      for (const { pattern, standard } of indicatorPatterns) {
        let match;
        const regex = new RegExp(pattern);
        
        while ((match = regex.exec(pageRange)) !== null) {
          if (standard === 'GRI' || standard === 'SASB') {
            extractedIndicators.push({
              reportId,
              chapterId: chapter._id,
              standardCode: `${standard} ${match[1]}`,
              category: match[2].trim(),
              name: match[2].trim(),
              page: estimatePageNumber(match.index, pageRange, chapter.startPage),
              context: extractContext(pageRange, match.index, 200)
            });
          } else {
            // 處理自定義數值模式
            const name = match[1].trim();
            const value = match[2].trim();
            const unit = match[3].trim();
            
            // 過濾掉明顯無關的數值（例如頁碼、章節編號等）
            if (value.length > 0 && name.length > 5 && !name.toLowerCase().includes('page') && !name.toLowerCase().includes('章')) {
              extractedIndicators.push({
                reportId,
                chapterId: chapter._id,
                standardCode: 'Custom',
                category: determineCategory(name),
                name,
                value,
                unit,
                page: estimatePageNumber(match.index, pageRange, chapter.startPage),
                context: extractContext(pageRange, match.index, 200)
              });
            }
          }
        }
      }
    }
    
    // 去除重複項
    const uniqueIndicators = removeDuplicateIndicators(extractedIndicators);
    
    // 將提取的指標保存到數據庫
    for (const indicator of uniqueIndicators) {
      const newIndicator = new Indicator(indicator);
      await newIndicator.save();
    }
    
    console.log(`從報告 ${reportId} 中提取了 ${uniqueIndicators.length} 個指標`);
    return uniqueIndicators.length;
  } catch (error) {
    console.error(`提取指標時出錯 reportId: ${reportId}:`, error);
    return 0;
  }
}

// 根據指標名稱確定類別
function determineCategory(name) {
  const name_lower = name.toLowerCase();
  
  // 環境類別
  if (name_lower.includes('排放') || name_lower.includes('碳') || name_lower.includes('能源') || 
      name_lower.includes('水') || name_lower.includes('廢棄物') || name_lower.includes('emission') ||
      name_lower.includes('carbon') || name_lower.includes('energy') || name_lower.includes('water') ||
      name_lower.includes('waste')) {
    return '環境指標';
  }
  
  // 社會類別
  if (name_lower.includes('員工') || name_lower.includes('勞工') || name_lower.includes('訓練') || 
      name_lower.includes('安全') || name_lower.includes('社區') || name_lower.includes('employee') ||
      name_lower.includes('labor') || name_lower.includes('training') || name_lower.includes('safety') ||
      name_lower.includes('community')) {
    return '社會指標';
  }
  
  // 治理類別
  if (name_lower.includes('治理') || name_lower.includes('董事') || name_lower.includes('風險') || 
      name_lower.includes('合規') || name_lower.includes('governance') || name_lower.includes('board') ||
      name_lower.includes('risk') || name_lower.includes('compliance')) {
    return '治理指標';
  }
  
  // 預設類別
  return '其他指標';
}

// 估計頁碼
function estimatePageNumber(matchIndex, text, startPage) {
  // 假設每頁平均有2000個字符
  const charsPerPage = 2000;
  const pageOffset = Math.floor(matchIndex / charsPerPage);
  return startPage + pageOffset;
}

// 提取上下文
function extractContext(text, position, contextLength) {
  const start = Math.max(0, position - contextLength / 2);
  const end = Math.min(text.length, position + contextLength / 2);
  return text.substring(start, end).replace(/\s+/g, ' ').trim();
}

// 去除重複指標
function removeDuplicateIndicators(indicators) {
  const uniqueMap = new Map();
  
  for (const indicator of indicators) {
    const key = `${indicator.name}-${indicator.value || ''}`;
    
    // 如果是新的指標或者比已有的指標有更完整的信息，則更新
    if (!uniqueMap.has(key) || 
        (indicator.value && !uniqueMap.get(key).value) ||
        (indicator.unit && !uniqueMap.get(key).unit)) {
      uniqueMap.set(key, indicator);
    }
  }
  
  return Array.from(uniqueMap.values());
}

// 提取永續目標
async function extractSustainabilityGoals(reportId, pdfBuffer) {
  try {
    // 獲取報告全文
    const data = await pdf(pdfBuffer);
    const fullText = data.text;
    
    // 獲取報告關聯的公司ID
    const report = await Report.findById(reportId);
    const companyId = report.companyId;
    
    // 常見目標模式 (例如: "到2030年減少50%的溫室氣體排放" 或 "2025年達成100%再生能源使用")
    const goalPatterns = [
      /(?:到|至|在|by)\s*(\d{4})(?:年|year)(?:前|底|末|end)?[，,\s]*([^\n\.。]{5,100})/g,
      /(\d{4})(?:年|year)(?:前|底|末|end)?[，,\s]*([^\n\.。]{5,100})/g,
      /目標(?:為|是|:)[，,\s]*([^\n\.。]{5,100})/g
    ];
    
    const extractedGoals = [];
    let match;
    
    // 應用各種模式提取目標
    for (const pattern of goalPatterns) {
      while ((match = pattern.exec(fullText)) !== null) {
        let targetYear, description;
        
        if (pattern.toString().includes('目標')) {
          // 處理沒有明確年份的目標
          description = match[1].trim();
          
          // 嘗試從描述中提取年份
          const yearMatch = description.match(/(\d{4})/);
          targetYear = yearMatch ? parseInt(yearMatch[1]) : null;
        } else {
          // 處理有明確年份的目標
          targetYear = parseInt(match[1]);
          description = match[2].trim();
        }
        
        // 確保描述不是太短或太長
        if (description.length > 10 && description.length < 200) {
          // 確定目標類別
          const category = determineGoalCategory(description);
          
          // 估計頁碼
          const page = estimatePageNumber(match.index, fullText, 1);
          
          extractedGoals.push({
            companyId,
            reportId,
            category,
            title: description.substring(0, Math.min(50, description.length)),
            description,
            targetYear,
            page,
            status: '進行中' // 預設狀態
          });
        }
      }
    }
    
    // 去除重複項
    const uniqueGoals = removeDuplicateGoals(extractedGoals);
    
    // 將提取的目標保存到數據庫
    for (const goal of uniqueGoals) {
      const newGoal = new Goal(goal);
      await newGoal.save();
    }
    
    console.log(`從報告 ${reportId} 中提取了 ${uniqueGoals.length} 個永續目標`);
    return uniqueGoals.length;
  } catch (error) {
    console.error(`提取永續目標時出錯 reportId: ${reportId}:`, error);
    return 0;
  }
}

// 確定目標類別
function determineGoalCategory(description) {
  const desc_lower = description.toLowerCase();
  
  // 減碳/氣候目標
  if (desc_lower.includes('碳') || desc_lower.includes('溫室氣體') || desc_lower.includes('排放') ||
      desc_lower.includes('淨零') || desc_lower.includes('carbon') || desc_lower.includes('emission') ||
      desc_lower.includes('net zero') || desc_lower.includes('ghg')) {
    return '減碳目標';
  }
  
  // 能源目標
  if (desc_lower.includes('能源') || desc_lower.includes('再生能源') || desc_lower.includes('energy') ||
      desc_lower.includes('renewable')) {
    return '能源目標';
  }
  
  // 水資源目標
  if (desc_lower.includes('水') || desc_lower.includes('water')) {
    return '水資源目標';
  }
  
  // 廢棄物目標
  if (desc_lower.includes('廢棄物') || desc_lower.includes('循環') || desc_lower.includes('waste') ||
      desc_lower.includes('circular')) {
    return '廢棄物目標';
  }
  
  // 社會/人才目標
  if (desc_lower.includes('員工') || desc_lower.includes('人才') || desc_lower.includes('多元') ||
      desc_lower.includes('包容') || desc_lower.includes('employee') || desc_lower.includes('talent') ||
      desc_lower.includes('diversity') || desc_lower.includes('inclusion')) {
    return '人才目標';
  }
  
  // 預設類別
  return '永續目標';
}

// 去除重複目標
function removeDuplicateGoals(goals) {
  const uniqueMap = new Map();
  
  for (const goal of goals) {
    // 使用描述的前50個字符作為識別鍵
    const key = goal.description.substring(0, 50);
    
    // 如果是新的目標或者比已有的目標有更完整的信息，則更新
    if (!uniqueMap.has(key) || 
        (goal.targetYear && !uniqueMap.get(key).targetYear)) {
      uniqueMap.set(key, goal);
    }
  }
  
  return Array.from(uniqueMap.values());
}

// 主要處理函數（暴露為 API）
module.exports = {
  processReport
};
