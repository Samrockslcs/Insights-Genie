/***** ================= Insights Geine – Full Restore + Resilience Logic ================= *****/

function doGet(e) {
  if (e && e.parameter && e.parameter.view) {
    try {
      const fileId = e.parameter.view;
      const file = DriveApp.getFileById(fileId);
      const htmlContent = file.getBlob().getDataAsString();

      return HtmlService.createHtmlOutput(htmlContent)
        .setTitle('Insights Geine Infographic')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (err) {
      return HtmlService.createHtmlOutput(
        "<h3>Error: Infographic not found or access denied.</h3>" +
        "<p>Please ensure the file exists and you have authorized Drive access.</p>"
      );
    }
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Insights Geine')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/***** ================= CONFIG ================= *****/
const GEMINI_KEY_PROP = 'GEMINI_API_KEY';
const GEMINI_MODEL = 'gemini-2.5-pro';
const GEMINI_BASE = 'https://generativelanguage.googleapis.com/v1beta/models/';
const DRIVE_FOLDER_ID = '1bB_2w_Dv36bIN3Kxr7WFNOcocicSz5A9';
const INFO_URL_PROP = 'Insights_Geine_LAST_INFO_URL';

/***** ================= CORE: Generate AI Insights ================= *****/
function generateInsightsFromSheet(sheetUrl, targetaudience) {
 try {
    const sheetId = extractSheetId_(sheetUrl);
    if (!sheetId) throw new Error('Invalid Sheet URL');

    const ss = SpreadsheetApp.openById(sheetId);
    
    // --- 1. PROCESS TERMS DATA (First Tab) ---
    const termSheet = ss.getSheets()[0];
    const termData = termSheet.getDataRange().getValues();
    if (termData.length < 2) throw new Error('No terms data found.');

    const tHeaders = termData[0].map(h => h.toString().toLowerCase());
    const qI = tHeaders.findIndex(h => h.includes('query'));
    const oI = tHeaders.findIndex(h => h.includes('opportunities') && !h.includes('pop'));
    const pI = tHeaders.findIndex(h => h.includes('pop'));

    if (qI === -1 || oI === -1 || pI === -1) throw new Error('Missing headers in Terms sheet: query, opportunity, PoP');

    const rows = termData.slice(1).map(r => ({
      query: String(r[qI] || '').trim(),
      adOpp: parseFloat(r[oI]) || 0,
      // FIX: Multiply by 100 because Sheets returns percentages as decimals (e.g., 0.5 for 50%)
      adOppPoP: (parseFloat(r[pI]) || 0) * 100 
    })).filter(r => r.query);

    const maxOpp = Math.max(...rows.map(r => r.adOpp));
    
    // Create the TERMS bundle
    const termEnriched = rows.map(r => ({
      query: r.query,
      IOS: maxOpp ? (r.adOpp / maxOpp) * 100 : 0,
      PoP: r.adOppPoP
    })).sort((a,b) => b.IOS - a.IOS);

    const termBundle = { top: termEnriched.slice(0, 100), totalRows: termEnriched.length };

    // --- 2. PROCESS LOCATION DATA (Tab: "Location_Data") ---
    let locationBundle = [];
    const locSheet = ss.getSheetByName("Location_Data");
    
    if (locSheet) {
      const locData = locSheet.getDataRange().getValues();
      if (locData.length > 1) {
        const lHeaders = locData[0].map(h => h.toString().toLowerCase());
        const gI = lHeaders.findIndex(h => h.includes('geo') || h.includes('name'));
        const loI = lHeaders.findIndex(h => h.includes('opportunities') && !h.includes('pop'));
        const lpI = lHeaders.findIndex(h => h.includes('pop'));

        if (gI > -1 && loI > -1 && lpI > -1) {
          locationBundle = locData.slice(1).map(r => ({
            geo: String(r[gI]),
            volume: parseFloat(r[loI]) || 0,
            // FIX: Multiply Location Growth by 100 as well
            growth: (parseFloat(r[lpI]) || 0) * 100
          }));
        }
      }
    }


    // --- 3. PROCESS INTENT DATA (Tab: "Intent" or "Intent_Data") ---
    let intentBundle = [];
    const intentSheet = ss.getSheetByName("Intent") || ss.getSheetByName("Intent_Data");
    if (intentSheet) {
      const intentData = intentSheet.getDataRange().getValues();
      if (intentData.length > 1) {
        const iHeaders = intentData[0].map(h => h.toString().toLowerCase());
        const iName = iHeaders.findIndex(h => h.includes('intent') || h.includes('query') || h.includes('name'));
        const iOpp = iHeaders.findIndex(h => h.includes('opportunities') && !h.includes('pop'));
        const iPop = iHeaders.findIndex(h => h.includes('pop'));
        
        intentBundle = intentData.slice(1).map(r => ({
          intent: String(r[iName > -1 ? iName : 0]),
          volume: parseFloat(r[iOpp > -1 ? iOpp : 1]) || 0,
          growth: (parseFloat(r[iPop > -1 ? iPop : 2]) || 0) * 100
        })).filter(r => r.intent);
      }
    }

    // --- 4. PROCESS FEATURES DATA (Tab: "Features" or "Features_Data") ---
    let featureBundle = [];
    const featureSheet = ss.getSheetByName("Features") || ss.getSheetByName("Features_Data");
    if (featureSheet) {
      const featureData = featureSheet.getDataRange().getValues();
      if (featureData.length > 1) {
        const fHeaders = featureData[0].map(h => h.toString().toLowerCase());
        const fName = fHeaders.findIndex(h => h.includes('feature') || h.includes('query') || h.includes('name'));
        const fOpp = fHeaders.findIndex(h => h.includes('opportunities') && !h.includes('pop'));
        const fPop = fHeaders.findIndex(h => h.includes('pop'));
        
        featureBundle = featureData.slice(1).map(r => ({
          feature: String(r[fName > -1 ? fName : 0]),
          volume: parseFloat(r[fOpp > -1 ? fOpp : 1]) || 0,
          growth: (parseFloat(r[fPop > -1 ? fPop : 2]) || 0) * 100
        })).filter(r => r.feature);
      }
    }

    // --- 5. PROCESS BRANDS DATA (Tab: "Brands" or "Brands_Data") ---
    let brandBundle = [];
    const brandSheet = ss.getSheetByName("Brands") || ss.getSheetByName("Brands_Data");
    if (brandSheet) {
      const brandData = brandSheet.getDataRange().getValues();
      if (brandData.length > 1) {
        const bHeaders = brandData[0].map(h => h.toString().toLowerCase());
        const bName = bHeaders.findIndex(h => h.includes('brand') || h.includes('query') || h.includes('name'));
        const bOpp = bHeaders.findIndex(h => h.includes('opportunities') && !h.includes('pop'));
        const bPop = bHeaders.findIndex(h => h.includes('pop'));
        
        brandBundle = brandData.slice(1).map(r => ({
          brand: String(r[bName > -1 ? bName : 0]),
          volume: parseFloat(r[bOpp > -1 ? bOpp : 1]) || 0,
          growth: (parseFloat(r[bPop > -1 ? bPop : 2]) || 0) * 100
        })).filter(r => r.brand);
      }
    }


    const prompt = `
ACT AS A SENIOR MARKET RESEARCH LEAD AND GOOGLE STRATEGIST. Create a high-impact, structured insight report based on the provided search data.
TOPIC: ${targetaudience}

DATA DEFINITIONS:
Topic: The keyword, feature, intent, or brand used to find this product.
Searches (Ad Opps): The raw volume of interest. CRITICAL FILTER: IGNORE any query with fewer than 1,000 Searches. Do not include these in any analysis.
YoY Growth: The percentage change in interest over the last year.

RAW DATASET (General Terms):
(Top 100 of ${termBundle.totalRows} queries):
${JSON.stringify(termBundle.top, null, 2)}

LOCATION DATA (Regional Breakdown):
${JSON.stringify(locationBundle, null, 2)}

INTENT DATA:
${JSON.stringify(intentBundle, null, 2)}

FEATURES DATA:
${JSON.stringify(featureBundle, null, 2)}

BRANDS DATA:
${JSON.stringify(brandBundle, null, 2)}

REPORT STRUCTURE:

1. TOP 3 EMERGING MARKET TRENDS
Identify overarching themes in the data (using only terms with >1,000 searches). For each trend, provide:
Trend Name: A descriptive title for the shift.
Metrics: Use [YoY Growth %] to justify the trend.
Evidence: Cite specific queries from the data that prove this trend exists for ${targetaudience}.

2. Brand Share of Searches
Using the search data, identify the presence of specific brands versus generic terms.
Relative Share Calculation: Do NOT use raw volumes. Instead, calculate the "Relative Share" percentage within the competitive set found in the data (e.g., if Brand A, B, and C are found, Brand A share = Brand A Volume / Total Volume of Brands A+B+C).
Analysis: Highlight which brands are dominating the Share of Mind for ${targetaudience} and which are losing ground based on YoY growth.
Format/Output: Present a clean breakdown showing the Calculated Share of Search (%) and the YoY Growth (%) for each brand. Add a brief qualitative summary of the competitive landscape (e.g., Who is the undisputed leader? Who are the rising challengers?).

3. BRAND X ATTRIBUTE
STRICT DATA ISOLATION RULE: For this section, you are strictly forbidden from looking at the "Brands", "Intent", or "Features" data tabs. You MUST derive the Brand x Attribute resonance EXCLUSIVELY by analyzing the granular queries within the "Top terms" data.
Objective: Find granular queries in the top terms where a specific brand is searched alongside a specific feature or attribute (e.g., a high-volume query like "[Brand A] camera" vs "[Brand B] battery"). Provide strong qualitative commentary on which brands are successfully owning specific narratives or attributes in the minds of consumers.
Format: Identify 3-4 interesting intersections directly from the query strings. Explain the consumer perception based on the search volume and growth of those specific raw queries.

4. The Intent Landscape
Analyze the "Intent" data to map where consumers are in the funnel.
Objective: Calculate the Share of Searches for different consumer intents relative to the total intent volume using Ad Opportunities.
Format: Present the Intent Share (%) and YoY Growth (%). Add commentary on shifting funnel dynamics (e.g., Is there a massive rise in "Comparison" or "Refurbished" queries indicating price sensitivity? Are "Review" searches outpacing "Best" searches?).

5.  Feature Demand Breakdown
Analyze the "Features" data to show what hardware/software capabilities matter most.
Objective: Calculate the Share of Searches for specific smartphone features relative to the total feature volume using Ad Opportunities.Have only top 6 by Ad Opportunities volume - the rest you can put in others
Format: Present the Feature Share (%) and YoY Growth (%). Add qualitative insights on which features are now considered "table stakes" (high share, flat growth) versus "innovation drivers" (lower share, explosive growth).

6. Strategic Marketing Recommendations (Google Ads & Content)
Translate the above insights into actionable sales pitches for the client.
Objective: Provide 4-5 concrete marketing recommendations.
Format: Group recommendations by strategy.
Content Strategy: How should brands pivot their ad creatives, video narratives, and messaging pillars based on the Features and Intent data?
Google Ads Strategy: How should they structure their Google Ads campaigns? (e.g., Bidding aggressively on mid-funnel comparison terms via Search, using explosive feature trends as hooks in YouTube Shorts campaigns, or leveraging Performance Max to capture shifting brand allegiances).

IMPORTANT:
Do NOT start with meta sentences like "Of course..." or "Here is your report".
Start directly from section 1.
No conversational filler.
STRICTLY PLAIN TEXT: Do NOT use Markdown formatting (no asterisks like ** or *).
Do NOT use bolding syntax.
Use uppercase for main headings (e.g., "1. TOP 3 EMERGING MARKET TRENDS") to distinguish them without bolding symbols.
THRESHOLD RULE: STRICTLY ignore all data points with "Searches" < 1,000.
VOLUME RULE: Never output raw search volume numbers. Use "Relative Share %" for groups/competitors and "YoY Growth %" elsewhere.
Ensure the advice is hyper-specific to ${targetaudience}.
`;

    const text = callGeminiText_(prompt);
    return { ok: true, reportText: text };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

/***** ================= GEMINI TEXT ================= *****/
function callGeminiText_(prompt) {
  const key = PropertiesService.getScriptProperties().getProperty(GEMINI_KEY_PROP);
  const url = `${GEMINI_BASE}${GEMINI_MODEL}:generateContent?key=${key}`;
  const payload = {
    contents: [{ role: 'user', parts: [{ text: prompt }]}],
    generationConfig: { temperature: 0.3, maxOutputTokens: 50000 }
  };
  const resp = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true, timeout: 60000
  });
  const parsed = JSON.parse(resp.getContentText() || '{}');
  return parsed.candidates?.[0]?.content?.parts?.[0]?.text || '';
}

/***** ================= GENERATE INFOGRAPHIC HTML ================= *****/
function generateInfographicHtmlFromInsights(targetaudience, reportText) {
  try {
    const key = PropertiesService.getScriptProperties().getProperty(GEMINI_KEY_PROP);
    const url = `${GEMINI_BASE}${GEMINI_MODEL}:generateContent?key=${key}`;

    const prompt = `You are generating a polished, modern Tailwind infographic.

STRICT OUTPUT RULES:
- Output ONLY HTML (no markdown, no code fences).
- Do NOT include <html>, <head>, <body>, or <script>.
- Use Tailwind utility classes only.
- Keep text concise. No long paragraphs.
- Like in the below example always keep the footer in the end.
- Also add the some explanation regarding the metric 
- Make sure that the infographics should contain all the insights while generating it.
- Always add the variety of chart wherever it make sense for each sections for eg table chart,pie hart,histogram ,bar chart etc.
- Horizontal Scroller: Applied to the "Emerging Market Trends" (Section 1) so users can swipe/scroll left to right through the trends.
- Horizontal Bar Chart: Applied to the "Brand Share" (Section 2) for a clean competitive landscape view.
- Brand Resonance Matrix: Enhanced the "Brand x Attribute" (Section 3) into a more distinct grid format.
- Donut Chart: Applied to "The Intent Landscape" (Section 4).
- Radar Chart: Applied to the "Feature Demand Breakdown" (Section 5) to visualize how different hardware features stack up against each other.
- Action Cards: Upgraded the "Strategic Marketing Recommendations" (Section 6) to look like distinct, actionable cards.
- "We recommend cross-checking all insights before sending to clients." this line should always be at the footer of the infographics

Maintain the theme of the example below across all infographics, but use complete insights:
${reportText} which is generated to make infographics:

<!DOCTYPE html><html><head><script src="https://cdn.tailwindcss.com"></script></head><body style="background:#f8fafc; padding:24px;"><div style="max-width:1152px; margin:0 auto; background:white; border-radius:16px; padding:32px; box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1);"><div class="max-w-6xl mx-auto bg-white border border-slate-200 rounded-2xl shadow-sm p-6 md:p-8 font-sans text-slate-800">

    <header class="max-w-6xl mx-auto mb-10 text-center">
        <div class="inline-block bg-blue-600 text-white px-4 py-1 rounded-full text-sm font-semibold tracking-wide uppercase mb-3">Search Insights Report</div>
        <h1 class="text-4xl md:text-5xl font-extrabold text-slate-900 mb-4">The Smartphone Market</h1>
        <p class="text-lg text-slate-600 max-w-3xl mx-auto">An analysis of consumer search trends, brand dynamics, and feature demand shaping the future of mobile technology.</p>
    </header>

    <main class="max-w-6xl mx-auto space-y-12">

        <!-- Section 1: Emerging Market Trends -->
        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">1. Top 3 Emerging Market Trends</h2>
                    <p class="text-slate-400 mt-1">Consumer behavior is radically shifting towards future-planning and practical AI application.</p>
                </div>
                <svg class="w-10 h-10 text-blue-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 7h8m0 0v8m0-8l-8 8-4-4-6 6"></path></svg>
            </div>
            
            <div class="relative">
                <div class="flex overflow-x-auto snap-x snap-mandatory scrollbar-thin scrollbar-thumb-slate-300 scrollbar-track-slate-100 p-6 md:p-8 space-x-6">
                    
                    <!-- Trend 1 Card -->
                    <div class="snap-center flex-shrink-0 w-11/12 md:w-2/3 lg:w-1/2 bg-slate-50 p-6 rounded-xl border border-slate-200">
                        <h3 class="text-lg font-semibold text-slate-800 mb-2">Pre-Launch Hype Cycle Acceleration</h3>
                        <p class="text-slate-600 mb-4 text-sm">Consumers are now planning purchases years in advance. Searches for unreleased models show massive YoY growth, indicating a need to engage audiences long before a product launch.</p>
                        <div class="grid grid-cols-2 gap-4">
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"iphone 17"</div><div class="text-2xl font-bold text-green-600 mt-1">+1398%</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"oppo reno 15"</div><div class="text-2xl font-bold text-green-600 mt-1">+109k%</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"iphone 17 pro max"</div><div class="text-2xl font-bold text-green-600 mt-1">+1023%</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"vivo v60"</div><div class="text-2xl font-bold text-green-600 mt-1">+27k%</div></div>
                        </div>
                    </div>

                    <!-- Trend 2 Card -->
                    <div class="snap-center flex-shrink-0 w-11/12 md:w-2/3 lg:w-1/2 bg-slate-50 p-6 rounded-xl border border-slate-200">
                        <h3 class="text-lg font-semibold text-slate-800 mb-2">The AI Utility Shift</h3>
                        <p class="text-slate-600 mb-4 text-sm">The conversation has moved from "What is AI?" to "How do I use it?". General AI queries are declining, while specific, use-case-oriented searches are exploding.</p>
                        <div class="grid grid-cols-1 gap-4">
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"new galaxy ai phone"</div><div class="text-2xl font-bold text-green-600 mt-1">+17,732% YoY</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"how to use circle to search"</div><div class="text-2xl font-bold text-green-600 mt-1">+821% YoY</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"writing assist samsung"</div><div class="text-2xl font-bold text-green-600 mt-1">+6000% YoY</div></div>
                        </div>
                    </div>

                    <!-- Trend 3 Card -->
                    <div class="snap-center flex-shrink-0 w-11/12 md:w-2/3 lg:w-1/2 bg-slate-50 p-6 rounded-xl border border-slate-200">
                        <h3 class="text-lg font-semibold text-slate-800 mb-2">Demand for Hyper-Specific Innovation</h3>
                        <p class="text-slate-600 mb-4 text-sm">Users are now researching niche, technical features that have a tangible impact on daily use, signaling a more sophisticated and demanding consumer base.</p>
                        <div class="grid grid-cols-1 gap-4">
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"privacy display samsung"</div><div class="text-2xl font-bold text-green-600 mt-1">+314,600% YoY</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"what is gemini on my phone"</div><div class="text-2xl font-bold text-green-600 mt-1">+46,100% YoY</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"priority notification iphone"</div><div class="text-2xl font-bold text-green-600 mt-1">+1558% YoY</div></div>
                        </div>
                    </div>
                </div>
            </div>
        </section>

        <!-- Section 2: Brand Share -->
        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">2. Brand Share of Searches</h2>
                    <p class="text-slate-400 mt-1">A two-tiered market where leaders solidify their position while challengers show dynamic growth.</p>
                </div>
                <svg class="w-10 h-10 text-indigo-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"></path></svg>
            </div>
            
            <div class="p-6 md:p-8 grid grid-cols-1 lg:grid-cols-5 gap-8">
                <div class="lg:col-span-3">
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Calculated Share & YoY Growth</h3>
                    <div class="space-y-4 text-sm">
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Apple</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 27.31%;">27.3%</div></div><div class="w-20 text-right font-medium text-green-600">+13.34%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Samsung</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 18.23%;">18.2%</div></div><div class="w-20 text-right font-medium text-red-600">-4.08%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">vivo</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 13.46%;">13.5%</div></div><div class="w-20 text-right font-medium text-green-600">+18.60%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">OPPO</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 8.76%;">8.8%</div></div><div class="w-20 text-right font-medium text-green-600">+20.53%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Realme</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 6.99%;">7.0%</div></div><div class="w-20 text-right font-medium text-red-600">-2.63%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Redmi</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 4.87%;">4.9%</div></div><div class="w-20 text-right font-medium text-red-600">-25.52%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Motorola</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 4.27%;">4.3%</div></div><div class="w-20 text-right font-medium text-green-600">+12.89%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Google</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 2.34%;">2.3%</div></div><div class="w-20 text-right font-medium text-green-600">+27.61%</div></div>
                    </div>
                </div>
                <div class="lg:col-span-2">
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Market Summary</h3>
                    <div class="space-y-3 text-sm text-slate-600">
                        <p><span class="font-bold text-slate-800">Leaders:</span> Apple solidifies its #1 spot with strong growth. Samsung maintains a huge share but shows signs of erosion.</p>
                        <p><span class="font-bold text-slate-800">Challengers:</span> Vivo and OPPO are standout growers, bucking the negative trend seen by competitors like Redmi and OnePlus.</p>
                        <p><span class="font-bold text-slate-800">Movers:</span> Google's Pixel line is gaining traction with notable growth. Huawei is staging a powerful comeback (+79% YoY), signaling a resurgence in consumer interest.</p>
                    </div>
                </div>
            </div>
        </section>

        <!-- Section 3: Brand x Attribute -->
        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">3. Brand x Attribute Resonance</h2>
                    <p class="text-slate-400 mt-1">How consumers perceive and search for brand-specific innovations.</p>
                </div>
                <svg class="w-10 h-10 text-teal-400" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M9.568 3H5.25A2.25 2.25 0 003 5.25v4.318c0 .597.237 1.17.659 1.591l9.581 9.581c.699.699 1.78.872 2.607.33a18.095 18.095 0 005.223-5.223c.542-.827.369-1.908-.33-2.607L11.16 3.66A2.25 2.25 0 009.568 3z" /><path stroke-linecap="round" stroke-linejoin="round" d="M6 6h.008v.008H6V6z" /></svg>
            </div>
            
            <div class="p-6 md:p-8 grid grid-cols-1 lg:grid-cols-3 gap-6">
                <div class="bg-slate-50 p-6 rounded-xl border border-slate-200">
                    <h3 class="text-lg font-semibold text-slate-800 mb-2">Samsung Owns "Next-Gen AI Utility"</h3>
                    <p class="text-slate-600 text-sm">Samsung has successfully branded its AI suite. Consumers associate them with tangible features, searching for <span class="font-semibold text-teal-700">"circle to search samsung"</span> and <span class="font-semibold text-teal-700">"samsung ai phone"</span>, proving they own the narrative of practical, in-hand innovation.</p>
                </div>
                <div class="bg-slate-50 p-6 rounded-xl border border-slate-200">
                    <h3 class="text-lg font-semibold text-slate-800 mb-2">Apple's Narrative: "Future Hype & AI Compatibility"</h3>
                    <p class="text-slate-600 text-sm">Apple dominates future product hype (<span class="font-semibold text-teal-700">"iphone 17"</span>). Their AI narrative is focused on compatibility, with queries like <span class="font-semibold text-teal-700">"apple intelligence supported phones"</span>, indicating a user mindset focused on access rather than specific use-cases.</p>
                </div>
                <div class="bg-slate-50 p-6 rounded-xl border border-slate-200">
                    <h3 class="text-lg font-semibold text-slate-800 mb-2">The Rise of the Hyper-Informed Searcher</h3>
                    <p class="text-slate-600 text-sm">Users are making decisions before they search. Highly detailed, near-purchase queries like <span class="font-semibold text-teal-700">"samsung galaxy s24 fe 5g ai smartphone"</span> show that brands must match this specificity to capture high-intent customers.</p>
                </div>
            </div>
        </section>

        <!-- Section 4: Intent Landscape -->
        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">4. The Intent Landscape</h2>
                    <p class="text-slate-400 mt-1">Price sensitivity is paramount, with explosive growth in value-based acquisition models.</p>
                </div>
                <svg class="w-10 h-10 text-amber-400" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M21 21l-5.197-5.197m0 0A7.5 7.5 0 105.196 5.196a7.5 7.5 0 0010.607 10.607z" /></svg>
            </div>
            
            <div class="p-6 md:p-8 grid grid-cols-1 lg:grid-cols-2 gap-8 items-center">
                <div class="relative w-64 h-64 mx-auto">
                    <div class="absolute inset-0 rounded-full" style="background-image: conic-gradient(from 0deg, #f59e0b 0% 50.43%, #64748b 50.43% 64.07%, #fbbf24 64.07% 77.42%, #fcd34d 77.42% 86.43%, #fef3c7 86.43% 100%);"></div>
                    <div class="absolute inset-5 bg-white rounded-full flex flex-col items-center justify-center">
                        <div class="text-4xl font-bold text-slate-800">50.4%</div>
                        <div class="text-sm text-slate-500 font-semibold">Price Intent</div>
                    </div>
                </div>
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Funnel Dynamics</h3>
                    <p class="text-slate-600 text-sm mb-4">While "Price" searches dominate, the real story is the massive growth in action-oriented, value-seeking behaviors. The decline in "Best" and "Reviews" suggests consumers are tired of generic advice and are focusing on the financial aspect of their purchase.</p>
                    <div class="grid grid-cols-2 gap-x-4 gap-y-2 text-sm">
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-amber-500 mr-2"></span>Price: <span class="ml-auto font-medium text-green-600">+4.2%</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-slate-500 mr-2"></span>Best: <span class="ml-auto font-medium text-red-600">-11.9%</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-amber-400 mr-2"></span>Comparison: <span class="ml-auto font-medium text-green-600">+4.1%</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-slate-500 mr-2"></span>Reviews: <span class="ml-auto font-medium text-red-600">-5.6%</span></div>
                        <div class="flex items-center col-span-2 border-t pt-2 mt-2 border-slate-200"><span class="font-bold text-slate-800">Growth Drivers:</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-green-500 mr-2"></span>Deals: <span class="ml-auto font-bold text-green-600">+93%</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-green-500 mr-2"></span>Used: <span class="ml-auto font-bold text-green-600">+123%</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-green-500 mr-2"></span>Trade-In: <span class="ml-auto font-bold text-green-600">+141%</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-green-500 mr-2"></span>Contract: <span class="ml-auto font-bold text-green-600">+187%</span></div>
                    </div>
                </div>
            </div>
        </section>

        <!-- Section 5: Feature Demand -->
        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">5. Feature Demand Breakdown</h2>
                    <p class="text-slate-400 mt-1">5G is now a baseline expectation. The new battlegrounds are screen technology and productivity.</p>
                </div>
                <svg class="w-10 h-10 text-rose-400" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M10.5 6h9.75M10.5 6a1.5 1.5 0 11-3 0m3 0a1.5 1.5 0 10-3 0M3.75 6H7.5m3 12h9.75m-9.75 0a1.5 1.5 0 01-3 0m3 0a1.5 1.5 0 00-3 0m-3.75 0H7.5m9-6h3.75m-3.75 0a1.5 1.5 0 01-3 0m3 0a1.5 1.5 0 00-3 0m-9.75 0h9.75" /></svg>
            </div>
            
            <div class="p-6 md:p-8 grid grid-cols-1 lg:grid-cols-2 gap-8 items-center">
                <div class="relative w-80 h-80 mx-auto">
                    <svg viewBox="0 0 200 200" class="w-full h-full">
                        <!-- Grid Lines -->
                        <polygon points="100,20 175,55 175,145 100,180 25,145 25,55" fill="none" stroke="#e2e8f0" stroke-width="1"/>
                        <polygon points="100,40 160,67 160,133 100,160 40,133 40,67" fill="none" stroke="#e2e8f0" stroke-width="1"/>
                        <polygon points="100,60 145,80 145,120 100,140 55,120 55,80" fill="none" stroke="#e2e8f0" stroke-width="1"/>
                        <line x1="100" y1="20" x2="100" y2="180" stroke="#e2e8f0" stroke-width="1"/>
                        <line x1="25" y1="55" x2="175" y2="145" stroke="#e2e8f0" stroke-width="1"/>
                        <line x1="25" y1="145" x2="175" y2="55" stroke="#e2e8f0" stroke-width="1"/>
                        <!-- Data Polygon -->
                        <polygon points="100,104.14 116.3,130.65 107.99,142.01 100,132.53 86.66,133.34 85.82,113.34" fill="rgba(244, 63, 94, 0.2)" stroke="#f43f5e" stroke-width="2"/>
                        <!-- Labels -->
                        <text x="100" y="15" text-anchor="middle" font-size="10" fill="#475569" font-weight="600">5G (-18%)</text>
                        <text x="180" y="55" text-anchor="middle" font-size="10" fill="#475569" font-weight="600">Screen (+102%)</text>
                        <text x="180" y="150" text-anchor="middle" font-size="10" fill="#475569" font-weight="600">Camera (+9%)</text>
                        <text x="100" y="195" text-anchor="middle" font-size="10" fill="#475569" font-weight="600">Color (+58%)</text>
                        <text x="20" y="150" text-anchor="middle" font-size="10" fill="#475569" font-weight="600">Battery (+21%)</text>
                        <text x="20" y="55" text-anchor="middle" font-size="10" fill="#475569" font-weight="600">Storage (+3%)</text>
                    </svg>
                </div>
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Feature Insights</h3>
                    <div class="space-y-4 text-sm text-slate-600">
                        <div>
                            <div class="font-bold text-slate-800">Table Stakes</div>
                            <p>"5G" has high search share but is declining, meaning consumers now expect it as standard. "Camera," "Battery," and "Storage" are evergreen pillars that require continuous competitive messaging.</p>
                        </div>
                        <div>
                            <div class="font-bold text-slate-800">Innovation Drivers</div>
                            <p>"Screen" is the breakout star (+102% YoY), showing that advancements in refresh rates, brightness, and privacy are capturing significant interest. "Color" (+58% YoY) also remains a powerful and growing differentiator.</p>
                        </div>
                    </div>
                </div>
            </div>
        </section>

        <!-- Section 6: Recommendations -->
        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">6. Strategic Marketing Recommendations</h2>
                    <p class="text-slate-400 mt-1">Actionable strategies for Google Ads and Content to capitalize on market shifts.</p>
                </div>
                <svg class="w-10 h-10 text-cyan-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9.663 17h4.673M12 3v1m6.364 1.636l-.707.707M21 12h-1M4 12H3m3.343-5.657l-.707-.707m2.828 9.9a5 5 0 117.072 0l-.548.547A3.374 3.374 0 0014 18.469V19a2 2 0 11-4 0v-.531c0-.895-.356-1.754-.988-2.386l-.548-.547z"></path></svg>
            </div>
            
            <div class="p-6 md:p-8 grid grid-cols-1 md:grid-cols-2 gap-8">
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Content Strategy</h3>
                    <div class="space-y-4">
                        <div class="p-4 bg-slate-50 rounded-lg border border-slate-200">
                            <div class="font-bold text-slate-800">1. Pivot AI Messaging from "What" to "How"</div>
                            <p class="text-sm text-slate-600 mt-1">Create content focused on utility. Develop guides and Shorts like "5 Ways Galaxy AI Can Organize Your Week" to capture the high-growth "how-to" search trend.</p>
                        </div>
                        <div class="p-4 bg-slate-50 rounded-lg border border-slate-200">
                            <div class="font-bold text-slate-800">2. Build Content Hubs for Innovation Drivers</div>
                            <p class="text-sm text-slate-600 mt-1">Develop deep content around "Screen Technology." Go beyond specs with articles on "How Privacy Screens Work" to capture sophisticated users searching for high-growth features.</p>
                        </div>
                    </div>
                </div>
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Google Ads Strategy</h3>
                    <div class="space-y-4">
                        <div class="p-4 bg-slate-50 rounded-lg border border-slate-200">
                            <div class="font-bold text-slate-800">3. Restructure Campaigns Around New Acquisition Models</div>
                            <p class="text-sm text-slate-600 mt-1">Build dedicated campaigns for "Contract," "Trade-In," and "Used" keywords. Lead ad copy with value: "Upgrade with a $400 Trade-In Bonus" to address market price sensitivity.</p>
                        </div>
                        <div class="p-4 bg-slate-50 rounded-lg border border-slate-200">
                            <div class="font-bold text-slate-800">4. Capture Switchers with Feature-Led PMax</div>
                            <p class="text-sm text-slate-600 mt-1">Use Performance Max to target "Comparison" searches. Use competitor brand names as audience signals and showcase high-growth features like vibrant "Colors" and "Screen" quality.</p>
                        </div>
                        <div class="p-4 bg-slate-50 rounded-lg border border-slate-200">
                            <div class="font-bold text-slate-800">5. Build Future Audiences with Pre-Launch Campaigns</div>
                            <p class="text-sm text-slate-600 mt-1">Capitalize on searches for "iPhone 17." Run light-touch YouTube/Discovery campaigns to build remarketing lists, then retarget them with pre-order messaging closer to launch.</p>
                        </div>
                    </div>
                </div>
            </div>
        </section>

    </main>

    <footer class="text-center text-sm text-slate-500 mt-12 pt-8 border-t border-slate-200">
        <p class="mt-2 text-slate-400 font-medium">We recommend cross-checking all insights before sending to clients.</p>
    </footer>

</div></div></body></html>

TITLE: ${targetaudience}

`;
  const payload = {
      contents: [{ role: 'user', parts: [{ text: prompt }]}],
      // Lowered maxOutputTokens slightly to prevent API truncation errors
      generationConfig: { temperature: 0.2, maxOutputTokens: 30000 }
    };

    const resp = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload), muteHttpExceptions: true, timeout: 60000
    });

    const responseText = resp.getContentText();
    let parsed = {};
    
    // Safely parse JSON in case the API throws an HTML error page (like a 502)
    try {
      parsed = JSON.parse(responseText);
    } catch (e) {
      throw new Error("Invalid API response format. The Gemini service might be overloaded.");
    }

    // Check for explicit API errors
    if (parsed.error) {
      throw new Error(`Gemini API Error: ${parsed.error.message}`);
    }

    // Extract content safely
    let htmlContent = parsed?.candidates?.[0]?.content?.parts?.[0]?.text || '';
    
    if (!htmlContent) {
      // Check if it was blocked by safety settings
      if (parsed?.candidates?.[0]?.finishReason === "SAFETY") {
        throw new Error("Content blocked by Google Safety Filters.");
      }
      throw new Error("Gemini returned an empty response. Try again.");
    }

    htmlContent = htmlContent.replace(/```html/gi, '').replace(/```/g, '').trim();

    // Pass to the Drive saver function
    const hostingResult = saveHtmlAndGetUrl_(targetaudience, htmlContent);
    return { ok: true, html: htmlContent, url: hostingResult.url };
    
  } catch (err) {
    // Send the EXACT error back to the frontend
    return { ok: false, error: err.message };
  }
}

/***** ================= HELPER: HOST ON DRIVE (REPAIRED) ================= *****/
function saveHtmlAndGetUrl_(label, htmlBody) {
  const fullHtml = `<!DOCTYPE html><html><head><script src="https://cdn.tailwindcss.com"></script></head><body style="background:#f8fafc; padding:24px;"><div style="max-width:1152px; margin:0 auto; background:white; border-radius:16px; padding:32px; box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1);">${htmlBody}</div></body></html>`;
  
  let folder;
  try {
    folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  } catch (e) {
    folder = DriveApp.getRootFolder();
  }
  
  const fileName = `Infographic-${label}-${Date.now()}.html`;
  const file = folder.createFile(fileName, fullHtml, MimeType.HTML);
  
  try {
    // Attempt 1: Domain-wide sharing (Best for corporate Workspaces)
    file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e) {
    try {
      // Attempt 2: Public Link fallback
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e2) {
      console.warn("Could not set sharing permissions. File remains private to the creator.");
    }
  }

  const scriptUrl = ScriptApp.getService().getUrl();
  // Ensure we format the URL correctly
  const finalUrl = scriptUrl ? `${scriptUrl}?view=${file.getId()}` : file.getUrl();

  return { url: finalUrl, fileId: file.getId() };
}

/***** ================= PUBLISH & EMAIL LOGIC ================= *****/
const LOGGING_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1ZfsNLTHFNitPQLxU1s0PHlR3y6z_DPtluldcZTEjyeU/edit';

function publishReportAndLog(audience, reportText, emailsCsv, person, client, rev) {
  try {
    const recipients = (emailsCsv || '').split(',').map(x => x.trim()).filter(Boolean);
    const userProps = PropertiesService.getUserProperties();
    const infoUrl = userProps.getProperty(INFO_URL_PROP) || '';

    // --- 1. Create Doc ---
    const doc = DocumentApp.create(`Insights Genie - ${audience} - ${new Date().toDateString()}`);
    const body = doc.getBody();
    body.appendParagraph(`AI Strategy Insights — ${audience}`).setHeading(DocumentApp.ParagraphHeading.TITLE);
    body.appendParagraph(`Strategist: ${person} | Client: ${client}`).setHeading(DocumentApp.ParagraphHeading.SUBTITLE);
    body.appendParagraph(reportText || '');

    if (infoUrl && infoUrl.startsWith('http')) {
      body.appendParagraph('Infographic').setHeading(DocumentApp.ParagraphHeading.HEADING2);
      const p = body.appendParagraph('📊 Open Infographic (Interactive)');
      p.setLinkUrl(infoUrl);
      p.setBold(true);
      p.setFontSize(13);
      p.setForegroundColor('#1a73e8');
    }
    doc.saveAndClose();

    // --- 2. Set Permissions ---
    const file = DriveApp.getFileById(doc.getId());
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    recipients.forEach(r => { try { file.addViewer(r); } catch (err) {} });

    // --- 3. Log to Tracking Sheet ---
    try {
      if (LOGGING_SHEET_URL) {
        const logSheet = SpreadsheetApp.openByUrl(LOGGING_SHEET_URL).getSheets()[0];
        logSheet.appendRow([
          new Date(),           // Timestamp
          person,               // Strategist Name
          client,               // Client Name
          audience,             // Topic
          doc.getUrl(),         // Link to the Google Doc
          infoUrl,              // Link to the Infographic
          recipients.join(', ') // Who it was sent to
        ]);
      }
    } catch (logErr) {
      console.log("Logging failed, but continuing execution: " + logErr);
    }

    // --- 4. Send Professional, No-Reply Email ---
    if (recipients.length > 0) {
      const emailSubject = `Strategic Insights Report: ${audience} | Prepared for ${client}`;
      
      const emailBody = `Dear Team,\n\n` +
        `Please find the strategic AI insights report for ${audience}, prepared specifically for ${client}.\n\n` +
        `This document contains a comprehensive analysis of the latest search trends, competitive dynamics, and actionable marketing recommendations based on consumer intent.\n\n` +
        `📄 Access the Strategy Document here: ${doc.getUrl()}\n\n` +
        `Best regards,\n` +
        `${person}`;

      // Using MailApp instead of GmailApp to enable 'noReply'
      MailApp.sendEmail({
        to: recipients.join(','),
        subject: emailSubject,
        body: emailBody,
        name: "Insights Genie", // Display name in the inbox
        noReply: true           // Forces a noreply@ address
      });
    }

    return { ok: true, docUrl: doc.getUrl(), url: infoUrl };
  } catch (err) {
    return { ok: false, error: "Publishing Error: " + err.message };
  }
}

/***** ================= UTILS ================= *****/
function extractSheetId_(url) {
  const m = String(url || '').match(/[-\w]{25,}/);
  return m ? m[0] : null;
}

function setInfographicHTML(html) { return true; }
function setInfographicURL(url) { 
  PropertiesService.getUserProperties().setProperty(INFO_URL_PROP, url);
  return true; 
}

/***** ================= USER CONTEXT (Via Admin Directory) ================= *****/
function getCurrentUserContext() {
  const email = Session.getActiveUser().getEmail();
  let fullName = "";
  let photoUrl = "";

  try {
    const user = AdminDirectory.Users.get(email, { viewType: 'domain_public' });
    if (user && user.name) {
      fullName = user.name.fullName || user.name.givenName;
    }
  } catch (e) {
    console.log("Admin Directory lookup failed: " + e.message);
  }

  if (!fullName) {
    const handle = email.split('@')[0];
    fullName = handle.split('.').map(part => part.charAt(0).toUpperCase() + part.slice(1)).join(' ');
  }

  return { email: email, name: fullName, photoUrl: photoUrl };
}

/***** ================= FORCE AUTHORIZATION ================= *****/
function authorizeScript() {
  DriveApp.getRootFolder(); 
  GmailApp.getInboxThreads(0, 1);
  
  // NEW: Forces Google to detect the MailApp permission requirement
  if (false) { MailApp.sendEmail("test@google.com", "test", "test"); }

  const tempDoc = DocumentApp.create("Insights Genie Permission Test");
  tempDoc.saveAndClose();
  DriveApp.getFileById(tempDoc.getId()).setTrashed(true);
  try {
    const email = Session.getActiveUser().getEmail();
    AdminDirectory.Users.get(email, { viewType: 'domain_public' });
  } catch (e) {
    console.warn('Admin Directory check failed: ' + e.message);
  }
}

/***** ================= CORE: Generate Custom Miscellaneous Insights ================= *****/
function generateMiscellaneousInsights(sheetUrl, targetaudience, customPrompt) {
  try {
    const sheetId = extractSheetId_(sheetUrl);
    if (!sheetId) throw new Error('Invalid Sheet URL');

    const ss = SpreadsheetApp.openById(sheetId);
    let miscSheet = ss.getSheetByName("Miscellaneous") || ss.getSheetByName("Misllenious");
    if (!miscSheet) throw new Error('Could not find a tab named "Miscellaneous" in the provided sheet.');

    const data = miscSheet.getDataRange().getDisplayValues();
    if (data.length < 2) throw new Error('No data found in the Miscellaneous tab.');

    const csvString = data.map(row => row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')).join('\n');

    const prompt = `ACT AS A SENIOR DATA ANALYST AND MARKET RESEARCHER.\nTOPIC: ${targetaudience}\nINSTRUCTIONS: ${customPrompt}\nDATA:\n${csvString}`;
    return { ok: true, reportText: callGeminiText_(prompt) };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function testMail() {
  MailApp.sendEmail(Session.getActiveUser().getEmail(), "Auth Test", "This is just to force the permission popup.");
}/***** ================= Insights Geine – Full Restore + Resilience Logic ================= *****/

function doGet(e) {
  if (e && e.parameter && e.parameter.view) {
    try {
      const fileId = e.parameter.view;
      const file = DriveApp.getFileById(fileId);
      const htmlContent = file.getBlob().getDataAsString();

      return HtmlService.createHtmlOutput(htmlContent)
        .setTitle('Insights Geine Infographic')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (err) {
      return HtmlService.createHtmlOutput(
        "<h3>Error: Infographic not found or access denied.</h3>" +
        "<p>Please ensure the file exists and you have authorized Drive access.</p>"
      );
    }
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Insights Geine')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/***** ================= CONFIG ================= *****/
const GEMINI_KEY_PROP = 'GEMINI_API_KEY';
const GEMINI_MODEL = 'gemini-2.5-pro';
const GEMINI_BASE = 'https://generativelanguage.googleapis.com/v1beta/models/';
const DRIVE_FOLDER_ID = '1bB_2w_Dv36bIN3Kxr7WFNOcocicSz5A9';
const INFO_URL_PROP = 'Insights_Geine_LAST_INFO_URL';

/***** ================= CORE: Generate AI Insights ================= *****/
function generateInsightsFromSheet(sheetUrl, targetaudience) {
 try {
    const sheetId = extractSheetId_(sheetUrl);
    if (!sheetId) throw new Error('Invalid Sheet URL');

    const ss = SpreadsheetApp.openById(sheetId);
    
    // --- 1. PROCESS TERMS DATA (First Tab) ---
    const termSheet = ss.getSheets()[0];
    const termData = termSheet.getDataRange().getValues();
    if (termData.length < 2) throw new Error('No terms data found.');

    const tHeaders = termData[0].map(h => h.toString().toLowerCase());
    const qI = tHeaders.findIndex(h => h.includes('query'));
    const oI = tHeaders.findIndex(h => h.includes('opportunities') && !h.includes('pop'));
    const pI = tHeaders.findIndex(h => h.includes('pop'));

    if (qI === -1 || oI === -1 || pI === -1) throw new Error('Missing headers in Terms sheet: query, opportunity, PoP');

    const rows = termData.slice(1).map(r => ({
      query: String(r[qI] || '').trim(),
      adOpp: parseFloat(r[oI]) || 0,
      // FIX: Multiply by 100 because Sheets returns percentages as decimals (e.g., 0.5 for 50%)
      adOppPoP: (parseFloat(r[pI]) || 0) * 100 
    })).filter(r => r.query);

    const maxOpp = Math.max(...rows.map(r => r.adOpp));
    
    // Create the TERMS bundle
    const termEnriched = rows.map(r => ({
      query: r.query,
      IOS: maxOpp ? (r.adOpp / maxOpp) * 100 : 0,
      PoP: r.adOppPoP
    })).sort((a,b) => b.IOS - a.IOS);

    const termBundle = { top: termEnriched.slice(0, 100), totalRows: termEnriched.length };

    // --- 2. PROCESS LOCATION DATA (Tab: "Location_Data") ---
    let locationBundle = [];
    const locSheet = ss.getSheetByName("Location_Data");
    
    if (locSheet) {
      const locData = locSheet.getDataRange().getValues();
      if (locData.length > 1) {
        const lHeaders = locData[0].map(h => h.toString().toLowerCase());
        const gI = lHeaders.findIndex(h => h.includes('geo') || h.includes('name'));
        const loI = lHeaders.findIndex(h => h.includes('opportunities') && !h.includes('pop'));
        const lpI = lHeaders.findIndex(h => h.includes('pop'));

        if (gI > -1 && loI > -1 && lpI > -1) {
          locationBundle = locData.slice(1).map(r => ({
            geo: String(r[gI]),
            volume: parseFloat(r[loI]) || 0,
            // FIX: Multiply Location Growth by 100 as well
            growth: (parseFloat(r[lpI]) || 0) * 100
          }));
        }
      }
    }


    // --- 3. PROCESS INTENT DATA (Tab: "Intent" or "Intent_Data") ---
    let intentBundle = [];
    const intentSheet = ss.getSheetByName("Intent") || ss.getSheetByName("Intent_Data");
    if (intentSheet) {
      const intentData = intentSheet.getDataRange().getValues();
      if (intentData.length > 1) {
        const iHeaders = intentData[0].map(h => h.toString().toLowerCase());
        const iName = iHeaders.findIndex(h => h.includes('intent') || h.includes('query') || h.includes('name'));
        const iOpp = iHeaders.findIndex(h => h.includes('opportunities') && !h.includes('pop'));
        const iPop = iHeaders.findIndex(h => h.includes('pop'));
        
        intentBundle = intentData.slice(1).map(r => ({
          intent: String(r[iName > -1 ? iName : 0]),
          volume: parseFloat(r[iOpp > -1 ? iOpp : 1]) || 0,
          growth: (parseFloat(r[iPop > -1 ? iPop : 2]) || 0) * 100
        })).filter(r => r.intent);
      }
    }

    // --- 4. PROCESS FEATURES DATA (Tab: "Features" or "Features_Data") ---
    let featureBundle = [];
    const featureSheet = ss.getSheetByName("Features") || ss.getSheetByName("Features_Data");
    if (featureSheet) {
      const featureData = featureSheet.getDataRange().getValues();
      if (featureData.length > 1) {
        const fHeaders = featureData[0].map(h => h.toString().toLowerCase());
        const fName = fHeaders.findIndex(h => h.includes('feature') || h.includes('query') || h.includes('name'));
        const fOpp = fHeaders.findIndex(h => h.includes('opportunities') && !h.includes('pop'));
        const fPop = fHeaders.findIndex(h => h.includes('pop'));
        
        featureBundle = featureData.slice(1).map(r => ({
          feature: String(r[fName > -1 ? fName : 0]),
          volume: parseFloat(r[fOpp > -1 ? fOpp : 1]) || 0,
          growth: (parseFloat(r[fPop > -1 ? fPop : 2]) || 0) * 100
        })).filter(r => r.feature);
      }
    }

    // --- 5. PROCESS BRANDS DATA (Tab: "Brands" or "Brands_Data") ---
    let brandBundle = [];
    const brandSheet = ss.getSheetByName("Brands") || ss.getSheetByName("Brands_Data");
    if (brandSheet) {
      const brandData = brandSheet.getDataRange().getValues();
      if (brandData.length > 1) {
        const bHeaders = brandData[0].map(h => h.toString().toLowerCase());
        const bName = bHeaders.findIndex(h => h.includes('brand') || h.includes('query') || h.includes('name'));
        const bOpp = bHeaders.findIndex(h => h.includes('opportunities') && !h.includes('pop'));
        const bPop = bHeaders.findIndex(h => h.includes('pop'));
        
        brandBundle = brandData.slice(1).map(r => ({
          brand: String(r[bName > -1 ? bName : 0]),
          volume: parseFloat(r[bOpp > -1 ? bOpp : 1]) || 0,
          growth: (parseFloat(r[bPop > -1 ? bPop : 2]) || 0) * 100
        })).filter(r => r.brand);
      }
    }


    const prompt = `
ACT AS A SENIOR MARKET RESEARCH LEAD AND GOOGLE STRATEGIST. Create a high-impact, structured insight report based on the provided search data.
TOPIC: ${targetaudience}

DATA DEFINITIONS:
Topic: The keyword, feature, intent, or brand used to find this product.
Searches (Ad Opps): The raw volume of interest. CRITICAL FILTER: IGNORE any query with fewer than 1,000 Searches. Do not include these in any analysis.
YoY Growth: The percentage change in interest over the last year.

RAW DATASET (General Terms):
(Top 100 of ${termBundle.totalRows} queries):
${JSON.stringify(termBundle.top, null, 2)}

LOCATION DATA (Regional Breakdown):
${JSON.stringify(locationBundle, null, 2)}

INTENT DATA:
${JSON.stringify(intentBundle, null, 2)}

FEATURES DATA:
${JSON.stringify(featureBundle, null, 2)}

BRANDS DATA:
${JSON.stringify(brandBundle, null, 2)}

REPORT STRUCTURE:

1. TOP 3 EMERGING MARKET TRENDS
Identify overarching themes in the data (using only terms with >1,000 searches). For each trend, provide:
Trend Name: A descriptive title for the shift.
Metrics: Use [YoY Growth %] to justify the trend.
Evidence: Cite specific queries from the data that prove this trend exists for ${targetaudience}.

2. Brand Share of Searches
Using the search data, identify the presence of specific brands versus generic terms.
Relative Share Calculation: Do NOT use raw volumes. Instead, calculate the "Relative Share" percentage within the competitive set found in the data (e.g., if Brand A, B, and C are found, Brand A share = Brand A Volume / Total Volume of Brands A+B+C).
Analysis: Highlight which brands are dominating the Share of Mind for ${targetaudience} and which are losing ground based on YoY growth.
Format/Output: Present a clean breakdown showing the Calculated Share of Search (%) and the YoY Growth (%) for each brand. Add a brief qualitative summary of the competitive landscape (e.g., Who is the undisputed leader? Who are the rising challengers?).

3. BRAND X ATTRIBUTE
STRICT DATA ISOLATION RULE: For this section, you are strictly forbidden from looking at the "Brands", "Intent", or "Features" data tabs. You MUST derive the Brand x Attribute resonance EXCLUSIVELY by analyzing the granular queries within the "Top terms" data.
Objective: Find granular queries in the top terms where a specific brand is searched alongside a specific feature or attribute (e.g., a high-volume query like "[Brand A] camera" vs "[Brand B] battery"). Provide strong qualitative commentary on which brands are successfully owning specific narratives or attributes in the minds of consumers.
Format: Identify 3-4 interesting intersections directly from the query strings. Explain the consumer perception based on the search volume and growth of those specific raw queries.

4. The Intent Landscape
Analyze the "Intent" data to map where consumers are in the funnel.
Objective: Calculate the Share of Searches for different consumer intents relative to the total intent volume using Ad Opportunities.
Format: Present the Intent Share (%) and YoY Growth (%). Add commentary on shifting funnel dynamics (e.g., Is there a massive rise in "Comparison" or "Refurbished" queries indicating price sensitivity? Are "Review" searches outpacing "Best" searches?).

5.  Feature Demand Breakdown
Analyze the "Features" data to show what hardware/software capabilities matter most.
Objective: Calculate the Share of Searches for specific smartphone features relative to the total feature volume using Ad Opportunities.Have only top 6 by Ad Opportunities volume - the rest you can put in others
Format: Present the Feature Share (%) and YoY Growth (%). Add qualitative insights on which features are now considered "table stakes" (high share, flat growth) versus "innovation drivers" (lower share, explosive growth).

6. Strategic Marketing Recommendations (Google Ads & Content)
Translate the above insights into actionable sales pitches for the client.
Objective: Provide 4-5 concrete marketing recommendations.
Format: Group recommendations by strategy.
Content Strategy: How should brands pivot their ad creatives, video narratives, and messaging pillars based on the Features and Intent data?
Google Ads Strategy: How should they structure their Google Ads campaigns? (e.g., Bidding aggressively on mid-funnel comparison terms via Search, using explosive feature trends as hooks in YouTube Shorts campaigns, or leveraging Performance Max to capture shifting brand allegiances).

IMPORTANT:
Do NOT start with meta sentences like "Of course..." or "Here is your report".
Start directly from section 1.
No conversational filler.
STRICTLY PLAIN TEXT: Do NOT use Markdown formatting (no asterisks like ** or *).
Do NOT use bolding syntax.
Use uppercase for main headings (e.g., "1. TOP 3 EMERGING MARKET TRENDS") to distinguish them without bolding symbols.
THRESHOLD RULE: STRICTLY ignore all data points with "Searches" < 1,000.
VOLUME RULE: Never output raw search volume numbers. Use "Relative Share %" for groups/competitors and "YoY Growth %" elsewhere.
Ensure the advice is hyper-specific to ${targetaudience}.
`;

    const text = callGeminiText_(prompt);
    return { ok: true, reportText: text };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

/***** ================= GEMINI TEXT ================= *****/
function callGeminiText_(prompt) {
  const key = PropertiesService.getScriptProperties().getProperty(GEMINI_KEY_PROP);
  const url = `${GEMINI_BASE}${GEMINI_MODEL}:generateContent?key=${key}`;
  const payload = {
    contents: [{ role: 'user', parts: [{ text: prompt }]}],
    generationConfig: { temperature: 0.3, maxOutputTokens: 50000 }
  };
  const resp = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true, timeout: 60000
  });
  const parsed = JSON.parse(resp.getContentText() || '{}');
  return parsed.candidates?.[0]?.content?.parts?.[0]?.text || '';
}

/***** ================= GENERATE INFOGRAPHIC HTML ================= *****/
function generateInfographicHtmlFromInsights(targetaudience, reportText) {
  try {
    const key = PropertiesService.getScriptProperties().getProperty(GEMINI_KEY_PROP);
    const url = `${GEMINI_BASE}${GEMINI_MODEL}:generateContent?key=${key}`;

    const prompt = `You are generating a polished, modern Tailwind infographic.

STRICT OUTPUT RULES:
- Output ONLY HTML (no markdown, no code fences).
- Do NOT include <html>, <head>, <body>, or <script>.
- Use Tailwind utility classes only.
- Keep text concise. No long paragraphs.
- Like in the below example always keep the footer in the end.
- Also add the some explanation regarding the metric 
- Make sure that the infographics should contain all the insights while generating it.
- Always add the variety of chart wherever it make sense for each sections for eg table chart,pie hart,histogram ,bar chart etc.
- Horizontal Scroller: Applied to the "Emerging Market Trends" (Section 1) so users can swipe/scroll left to right through the trends.
- Horizontal Bar Chart: Applied to the "Brand Share" (Section 2) for a clean competitive landscape view.
- Brand Resonance Matrix: Enhanced the "Brand x Attribute" (Section 3) into a more distinct grid format.
- Donut Chart: Applied to "The Intent Landscape" (Section 4).
- Radar Chart: Applied to the "Feature Demand Breakdown" (Section 5) to visualize how different hardware features stack up against each other.
- Action Cards: Upgraded the "Strategic Marketing Recommendations" (Section 6) to look like distinct, actionable cards.
- "We recommend cross-checking all insights before sending to clients." this line should always be at the footer of the infographics

Maintain the theme of the example below across all infographics, but use complete insights:
${reportText} which is generated to make infographics:

<!DOCTYPE html><html><head><script src="https://cdn.tailwindcss.com"></script></head><body style="background:#f8fafc; padding:24px;"><div style="max-width:1152px; margin:0 auto; background:white; border-radius:16px; padding:32px; box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1);"><div class="max-w-6xl mx-auto bg-white border border-slate-200 rounded-2xl shadow-sm p-6 md:p-8 font-sans text-slate-800">

    <header class="max-w-6xl mx-auto mb-10 text-center">
        <div class="inline-block bg-blue-600 text-white px-4 py-1 rounded-full text-sm font-semibold tracking-wide uppercase mb-3">Search Insights Report</div>
        <h1 class="text-4xl md:text-5xl font-extrabold text-slate-900 mb-4">The Smartphone Market</h1>
        <p class="text-lg text-slate-600 max-w-3xl mx-auto">An analysis of consumer search trends, brand dynamics, and feature demand shaping the future of mobile technology.</p>
    </header>

    <main class="max-w-6xl mx-auto space-y-12">

        <!-- Section 1: Emerging Market Trends -->
        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">1. Top 3 Emerging Market Trends</h2>
                    <p class="text-slate-400 mt-1">Consumer behavior is radically shifting towards future-planning and practical AI application.</p>
                </div>
                <svg class="w-10 h-10 text-blue-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 7h8m0 0v8m0-8l-8 8-4-4-6 6"></path></svg>
            </div>
            
            <div class="relative">
                <div class="flex overflow-x-auto snap-x snap-mandatory scrollbar-thin scrollbar-thumb-slate-300 scrollbar-track-slate-100 p-6 md:p-8 space-x-6">
                    
                    <!-- Trend 1 Card -->
                    <div class="snap-center flex-shrink-0 w-11/12 md:w-2/3 lg:w-1/2 bg-slate-50 p-6 rounded-xl border border-slate-200">
                        <h3 class="text-lg font-semibold text-slate-800 mb-2">Pre-Launch Hype Cycle Acceleration</h3>
                        <p class="text-slate-600 mb-4 text-sm">Consumers are now planning purchases years in advance. Searches for unreleased models show massive YoY growth, indicating a need to engage audiences long before a product launch.</p>
                        <div class="grid grid-cols-2 gap-4">
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"iphone 17"</div><div class="text-2xl font-bold text-green-600 mt-1">+1398%</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"oppo reno 15"</div><div class="text-2xl font-bold text-green-600 mt-1">+109k%</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"iphone 17 pro max"</div><div class="text-2xl font-bold text-green-600 mt-1">+1023%</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"vivo v60"</div><div class="text-2xl font-bold text-green-600 mt-1">+27k%</div></div>
                        </div>
                    </div>

                    <!-- Trend 2 Card -->
                    <div class="snap-center flex-shrink-0 w-11/12 md:w-2/3 lg:w-1/2 bg-slate-50 p-6 rounded-xl border border-slate-200">
                        <h3 class="text-lg font-semibold text-slate-800 mb-2">The AI Utility Shift</h3>
                        <p class="text-slate-600 mb-4 text-sm">The conversation has moved from "What is AI?" to "How do I use it?". General AI queries are declining, while specific, use-case-oriented searches are exploding.</p>
                        <div class="grid grid-cols-1 gap-4">
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"new galaxy ai phone"</div><div class="text-2xl font-bold text-green-600 mt-1">+17,732% YoY</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"how to use circle to search"</div><div class="text-2xl font-bold text-green-600 mt-1">+821% YoY</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"writing assist samsung"</div><div class="text-2xl font-bold text-green-600 mt-1">+6000% YoY</div></div>
                        </div>
                    </div>

                    <!-- Trend 3 Card -->
                    <div class="snap-center flex-shrink-0 w-11/12 md:w-2/3 lg:w-1/2 bg-slate-50 p-6 rounded-xl border border-slate-200">
                        <h3 class="text-lg font-semibold text-slate-800 mb-2">Demand for Hyper-Specific Innovation</h3>
                        <p class="text-slate-600 mb-4 text-sm">Users are now researching niche, technical features that have a tangible impact on daily use, signaling a more sophisticated and demanding consumer base.</p>
                        <div class="grid grid-cols-1 gap-4">
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"privacy display samsung"</div><div class="text-2xl font-bold text-green-600 mt-1">+314,600% YoY</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"what is gemini on my phone"</div><div class="text-2xl font-bold text-green-600 mt-1">+46,100% YoY</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"priority notification iphone"</div><div class="text-2xl font-bold text-green-600 mt-1">+1558% YoY</div></div>
                        </div>
                    </div>
                </div>
            </div>
        </section>

        <!-- Section 2: Brand Share -->
        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">2. Brand Share of Searches</h2>
                    <p class="text-slate-400 mt-1">A two-tiered market where leaders solidify their position while challengers show dynamic growth.</p>
                </div>
                <svg class="w-10 h-10 text-indigo-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"></path></svg>
            </div>
            
            <div class="p-6 md:p-8 grid grid-cols-1 lg:grid-cols-5 gap-8">
                <div class="lg:col-span-3">
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Calculated Share & YoY Growth</h3>
                    <div class="space-y-4 text-sm">
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Apple</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 27.31%;">27.3%</div></div><div class="w-20 text-right font-medium text-green-600">+13.34%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Samsung</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 18.23%;">18.2%</div></div><div class="w-20 text-right font-medium text-red-600">-4.08%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">vivo</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 13.46%;">13.5%</div></div><div class="w-20 text-right font-medium text-green-600">+18.60%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">OPPO</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 8.76%;">8.8%</div></div><div class="w-20 text-right font-medium text-green-600">+20.53%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Realme</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 6.99%;">7.0%</div></div><div class="w-20 text-right font-medium text-red-600">-2.63%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Redmi</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 4.87%;">4.9%</div></div><div class="w-20 text-right font-medium text-red-600">-25.52%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Motorola</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 4.27%;">4.3%</div></div><div class="w-20 text-right font-medium text-green-600">+12.89%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Google</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 2.34%;">2.3%</div></div><div class="w-20 text-right font-medium text-green-600">+27.61%</div></div>
                    </div>
                </div>
                <div class="lg:col-span-2">
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Market Summary</h3>
                    <div class="space-y-3 text-sm text-slate-600">
                        <p><span class="font-bold text-slate-800">Leaders:</span> Apple solidifies its #1 spot with strong growth. Samsung maintains a huge share but shows signs of erosion.</p>
                        <p><span class="font-bold text-slate-800">Challengers:</span> Vivo and OPPO are standout growers, bucking the negative trend seen by competitors like Redmi and OnePlus.</p>
                        <p><span class="font-bold text-slate-800">Movers:</span> Google's Pixel line is gaining traction with notable growth. Huawei is staging a powerful comeback (+79% YoY), signaling a resurgence in consumer interest.</p>
                    </div>
                </div>
            </div>
        </section>

        <!-- Section 3: Brand x Attribute -->
        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">3. Brand x Attribute Resonance</h2>
                    <p class="text-slate-400 mt-1">How consumers perceive and search for brand-specific innovations.</p>
                </div>
                <svg class="w-10 h-10 text-teal-400" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M9.568 3H5.25A2.25 2.25 0 003 5.25v4.318c0 .597.237 1.17.659 1.591l9.581 9.581c.699.699 1.78.872 2.607.33a18.095 18.095 0 005.223-5.223c.542-.827.369-1.908-.33-2.607L11.16 3.66A2.25 2.25 0 009.568 3z" /><path stroke-linecap="round" stroke-linejoin="round" d="M6 6h.008v.008H6V6z" /></svg>
            </div>
            
            <div class="p-6 md:p-8 grid grid-cols-1 lg:grid-cols-3 gap-6">
                <div class="bg-slate-50 p-6 rounded-xl border border-slate-200">
                    <h3 class="text-lg font-semibold text-slate-800 mb-2">Samsung Owns "Next-Gen AI Utility"</h3>
                    <p class="text-slate-600 text-sm">Samsung has successfully branded its AI suite. Consumers associate them with tangible features, searching for <span class="font-semibold text-teal-700">"circle to search samsung"</span> and <span class="font-semibold text-teal-700">"samsung ai phone"</span>, proving they own the narrative of practical, in-hand innovation.</p>
                </div>
                <div class="bg-slate-50 p-6 rounded-xl border border-slate-200">
                    <h3 class="text-lg font-semibold text-slate-800 mb-2">Apple's Narrative: "Future Hype & AI Compatibility"</h3>
                    <p class="text-slate-600 text-sm">Apple dominates future product hype (<span class="font-semibold text-teal-700">"iphone 17"</span>). Their AI narrative is focused on compatibility, with queries like <span class="font-semibold text-teal-700">"apple intelligence supported phones"</span>, indicating a user mindset focused on access rather than specific use-cases.</p>
                </div>
                <div class="bg-slate-50 p-6 rounded-xl border border-slate-200">
                    <h3 class="text-lg font-semibold text-slate-800 mb-2">The Rise of the Hyper-Informed Searcher</h3>
                    <p class="text-slate-600 text-sm">Users are making decisions before they search. Highly detailed, near-purchase queries like <span class="font-semibold text-teal-700">"samsung galaxy s24 fe 5g ai smartphone"</span> show that brands must match this specificity to capture high-intent customers.</p>
                </div>
            </div>
        </section>

        <!-- Section 4: Intent Landscape -->
        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">4. The Intent Landscape</h2>
                    <p class="text-slate-400 mt-1">Price sensitivity is paramount, with explosive growth in value-based acquisition models.</p>
                </div>
                <svg class="w-10 h-10 text-amber-400" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M21 21l-5.197-5.197m0 0A7.5 7.5 0 105.196 5.196a7.5 7.5 0 0010.607 10.607z" /></svg>
            </div>
            
            <div class="p-6 md:p-8 grid grid-cols-1 lg:grid-cols-2 gap-8 items-center">
                <div class="relative w-64 h-64 mx-auto">
                    <div class="absolute inset-0 rounded-full" style="background-image: conic-gradient(from 0deg, #f59e0b 0% 50.43%, #64748b 50.43% 64.07%, #fbbf24 64.07% 77.42%, #fcd34d 77.42% 86.43%, #fef3c7 86.43% 100%);"></div>
                    <div class="absolute inset-5 bg-white rounded-full flex flex-col items-center justify-center">
                        <div class="text-4xl font-bold text-slate-800">50.4%</div>
                        <div class="text-sm text-slate-500 font-semibold">Price Intent</div>
                    </div>
                </div>
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Funnel Dynamics</h3>
                    <p class="text-slate-600 text-sm mb-4">While "Price" searches dominate, the real story is the massive growth in action-oriented, value-seeking behaviors. The decline in "Best" and "Reviews" suggests consumers are tired of generic advice and are focusing on the financial aspect of their purchase.</p>
                    <div class="grid grid-cols-2 gap-x-4 gap-y-2 text-sm">
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-amber-500 mr-2"></span>Price: <span class="ml-auto font-medium text-green-600">+4.2%</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-slate-500 mr-2"></span>Best: <span class="ml-auto font-medium text-red-600">-11.9%</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-amber-400 mr-2"></span>Comparison: <span class="ml-auto font-medium text-green-600">+4.1%</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-slate-500 mr-2"></span>Reviews: <span class="ml-auto font-medium text-red-600">-5.6%</span></div>
                        <div class="flex items-center col-span-2 border-t pt-2 mt-2 border-slate-200"><span class="font-bold text-slate-800">Growth Drivers:</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-green-500 mr-2"></span>Deals: <span class="ml-auto font-bold text-green-600">+93%</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-green-500 mr-2"></span>Used: <span class="ml-auto font-bold text-green-600">+123%</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-green-500 mr-2"></span>Trade-In: <span class="ml-auto font-bold text-green-600">+141%</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-green-500 mr-2"></span>Contract: <span class="ml-auto font-bold text-green-600">+187%</span></div>
                    </div>
                </div>
            </div>
        </section>

        <!-- Section 5: Feature Demand -->
        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">5. Feature Demand Breakdown</h2>
                    <p class="text-slate-400 mt-1">5G is now a baseline expectation. The new battlegrounds are screen technology and productivity.</p>
                </div>
                <svg class="w-10 h-10 text-rose-400" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M10.5 6h9.75M10.5 6a1.5 1.5 0 11-3 0m3 0a1.5 1.5 0 10-3 0M3.75 6H7.5m3 12h9.75m-9.75 0a1.5 1.5 0 01-3 0m3 0a1.5 1.5 0 00-3 0m-3.75 0H7.5m9-6h3.75m-3.75 0a1.5 1.5 0 01-3 0m3 0a1.5 1.5 0 00-3 0m-9.75 0h9.75" /></svg>
            </div>
            
            <div class="p-6 md:p-8 grid grid-cols-1 lg:grid-cols-2 gap-8 items-center">
                <div class="relative w-80 h-80 mx-auto">
                    <svg viewBox="0 0 200 200" class="w-full h-full">
                        <!-- Grid Lines -->
                        <polygon points="100,20 175,55 175,145 100,180 25,145 25,55" fill="none" stroke="#e2e8f0" stroke-width="1"/>
                        <polygon points="100,40 160,67 160,133 100,160 40,133 40,67" fill="none" stroke="#e2e8f0" stroke-width="1"/>
                        <polygon points="100,60 145,80 145,120 100,140 55,120 55,80" fill="none" stroke="#e2e8f0" stroke-width="1"/>
                        <line x1="100" y1="20" x2="100" y2="180" stroke="#e2e8f0" stroke-width="1"/>
                        <line x1="25" y1="55" x2="175" y2="145" stroke="#e2e8f0" stroke-width="1"/>
                        <line x1="25" y1="145" x2="175" y2="55" stroke="#e2e8f0" stroke-width="1"/>
                        <!-- Data Polygon -->
                        <polygon points="100,104.14 116.3,130.65 107.99,142.01 100,132.53 86.66,133.34 85.82,113.34" fill="rgba(244, 63, 94, 0.2)" stroke="#f43f5e" stroke-width="2"/>
                        <!-- Labels -->
                        <text x="100" y="15" text-anchor="middle" font-size="10" fill="#475569" font-weight="600">5G (-18%)</text>
                        <text x="180" y="55" text-anchor="middle" font-size="10" fill="#475569" font-weight="600">Screen (+102%)</text>
                        <text x="180" y="150" text-anchor="middle" font-size="10" fill="#475569" font-weight="600">Camera (+9%)</text>
                        <text x="100" y="195" text-anchor="middle" font-size="10" fill="#475569" font-weight="600">Color (+58%)</text>
                        <text x="20" y="150" text-anchor="middle" font-size="10" fill="#475569" font-weight="600">Battery (+21%)</text>
                        <text x="20" y="55" text-anchor="middle" font-size="10" fill="#475569" font-weight="600">Storage (+3%)</text>
                    </svg>
                </div>
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Feature Insights</h3>
                    <div class="space-y-4 text-sm text-slate-600">
                        <div>
                            <div class="font-bold text-slate-800">Table Stakes</div>
                            <p>"5G" has high search share but is declining, meaning consumers now expect it as standard. "Camera," "Battery," and "Storage" are evergreen pillars that require continuous competitive messaging.</p>
                        </div>
                        <div>
                            <div class="font-bold text-slate-800">Innovation Drivers</div>
                            <p>"Screen" is the breakout star (+102% YoY), showing that advancements in refresh rates, brightness, and privacy are capturing significant interest. "Color" (+58% YoY) also remains a powerful and growing differentiator.</p>
                        </div>
                    </div>
                </div>
            </div>
        </section>

        <!-- Section 6: Recommendations -->
        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">6. Strategic Marketing Recommendations</h2>
                    <p class="text-slate-400 mt-1">Actionable strategies for Google Ads and Content to capitalize on market shifts.</p>
                </div>
                <svg class="w-10 h-10 text-cyan-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9.663 17h4.673M12 3v1m6.364 1.636l-.707.707M21 12h-1M4 12H3m3.343-5.657l-.707-.707m2.828 9.9a5 5 0 117.072 0l-.548.547A3.374 3.374 0 0014 18.469V19a2 2 0 11-4 0v-.531c0-.895-.356-1.754-.988-2.386l-.548-.547z"></path></svg>
            </div>
            
            <div class="p-6 md:p-8 grid grid-cols-1 md:grid-cols-2 gap-8">
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Content Strategy</h3>
                    <div class="space-y-4">
                        <div class="p-4 bg-slate-50 rounded-lg border border-slate-200">
                            <div class="font-bold text-slate-800">1. Pivot AI Messaging from "What" to "How"</div>
                            <p class="text-sm text-slate-600 mt-1">Create content focused on utility. Develop guides and Shorts like "5 Ways Galaxy AI Can Organize Your Week" to capture the high-growth "how-to" search trend.</p>
                        </div>
                        <div class="p-4 bg-slate-50 rounded-lg border border-slate-200">
                            <div class="font-bold text-slate-800">2. Build Content Hubs for Innovation Drivers</div>
                            <p class="text-sm text-slate-600 mt-1">Develop deep content around "Screen Technology." Go beyond specs with articles on "How Privacy Screens Work" to capture sophisticated users searching for high-growth features.</p>
                        </div>
                    </div>
                </div>
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Google Ads Strategy</h3>
                    <div class="space-y-4">
                        <div class="p-4 bg-slate-50 rounded-lg border border-slate-200">
                            <div class="font-bold text-slate-800">3. Restructure Campaigns Around New Acquisition Models</div>
                            <p class="text-sm text-slate-600 mt-1">Build dedicated campaigns for "Contract," "Trade-In," and "Used" keywords. Lead ad copy with value: "Upgrade with a $400 Trade-In Bonus" to address market price sensitivity.</p>
                        </div>
                        <div class="p-4 bg-slate-50 rounded-lg border border-slate-200">
                            <div class="font-bold text-slate-800">4. Capture Switchers with Feature-Led PMax</div>
                            <p class="text-sm text-slate-600 mt-1">Use Performance Max to target "Comparison" searches. Use competitor brand names as audience signals and showcase high-growth features like vibrant "Colors" and "Screen" quality.</p>
                        </div>
                        <div class="p-4 bg-slate-50 rounded-lg border border-slate-200">
                            <div class="font-bold text-slate-800">5. Build Future Audiences with Pre-Launch Campaigns</div>
                            <p class="text-sm text-slate-600 mt-1">Capitalize on searches for "iPhone 17." Run light-touch YouTube/Discovery campaigns to build remarketing lists, then retarget them with pre-order messaging closer to launch.</p>
                        </div>
                    </div>
                </div>
            </div>
        </section>

    </main>

    <footer class="text-center text-sm text-slate-500 mt-12 pt-8 border-t border-slate-200">
        <p class="mt-2 text-slate-400 font-medium">We recommend cross-checking all insights before sending to clients.</p>
    </footer>

</div></div></body></html>

TITLE: ${targetaudience}

`;
  const payload = {
      contents: [{ role: 'user', parts: [{ text: prompt }]}],
      // Lowered maxOutputTokens slightly to prevent API truncation errors
      generationConfig: { temperature: 0.2, maxOutputTokens: 30000 }
    };

    const resp = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload), muteHttpExceptions: true, timeout: 60000
    });

    const responseText = resp.getContentText();
    let parsed = {};
    
    // Safely parse JSON in case the API throws an HTML error page (like a 502)
    try {
      parsed = JSON.parse(responseText);
    } catch (e) {
      throw new Error("Invalid API response format. The Gemini service might be overloaded.");
    }

    // Check for explicit API errors
    if (parsed.error) {
      throw new Error(`Gemini API Error: ${parsed.error.message}`);
    }

    // Extract content safely
    let htmlContent = parsed?.candidates?.[0]?.content?.parts?.[0]?.text || '';
    
    if (!htmlContent) {
      // Check if it was blocked by safety settings
      if (parsed?.candidates?.[0]?.finishReason === "SAFETY") {
        throw new Error("Content blocked by Google Safety Filters.");
      }
      throw new Error("Gemini returned an empty response. Try again.");
    }

    htmlContent = htmlContent.replace(/```html/gi, '').replace(/```/g, '').trim();

    // Pass to the Drive saver function
    const hostingResult = saveHtmlAndGetUrl_(targetaudience, htmlContent);
    return { ok: true, html: htmlContent, url: hostingResult.url };
    
  } catch (err) {
    // Send the EXACT error back to the frontend
    return { ok: false, error: err.message };
  }
}

/***** ================= HELPER: HOST ON DRIVE (REPAIRED) ================= *****/
function saveHtmlAndGetUrl_(label, htmlBody) {
  const fullHtml = `<!DOCTYPE html><html><head><script src="https://cdn.tailwindcss.com"></script></head><body style="background:#f8fafc; padding:24px;"><div style="max-width:1152px; margin:0 auto; background:white; border-radius:16px; padding:32px; box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1);">${htmlBody}</div></body></html>`;
  
  let folder;
  try {
    folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  } catch (e) {
    folder = DriveApp.getRootFolder();
  }
  
  const fileName = `Infographic-${label}-${Date.now()}.html`;
  const file = folder.createFile(fileName, fullHtml, MimeType.HTML);
  
  try {
    // Attempt 1: Domain-wide sharing (Best for corporate Workspaces)
    file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e) {
    try {
      // Attempt 2: Public Link fallback
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e2) {
      console.warn("Could not set sharing permissions. File remains private to the creator.");
    }
  }

  const scriptUrl = ScriptApp.getService().getUrl();
  // Ensure we format the URL correctly
  const finalUrl = scriptUrl ? `${scriptUrl}?view=${file.getId()}` : file.getUrl();

  return { url: finalUrl, fileId: file.getId() };
}

/***** ================= PUBLISH & EMAIL LOGIC ================= *****/
const LOGGING_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1ZfsNLTHFNitPQLxU1s0PHlR3y6z_DPtluldcZTEjyeU/edit';

function publishReportAndLog(audience, reportText, emailsCsv, person, client, rev) {
  try {
    const recipients = (emailsCsv || '').split(',').map(x => x.trim()).filter(Boolean);
    const userProps = PropertiesService.getUserProperties();
    const infoUrl = userProps.getProperty(INFO_URL_PROP) || '';

    // --- 1. Create Doc ---
    const doc = DocumentApp.create(`Insights Genie - ${audience} - ${new Date().toDateString()}`);
    const body = doc.getBody();
    body.appendParagraph(`AI Strategy Insights — ${audience}`).setHeading(DocumentApp.ParagraphHeading.TITLE);
    body.appendParagraph(`Strategist: ${person} | Client: ${client}`).setHeading(DocumentApp.ParagraphHeading.SUBTITLE);
    body.appendParagraph(reportText || '');

    if (infoUrl && infoUrl.startsWith('http')) {
      body.appendParagraph('Infographic').setHeading(DocumentApp.ParagraphHeading.HEADING2);
      const p = body.appendParagraph('📊 Open Infographic (Interactive)');
      p.setLinkUrl(infoUrl);
      p.setBold(true);
      p.setFontSize(13);
      p.setForegroundColor('#1a73e8');
    }
    doc.saveAndClose();

    // --- 2. Set Permissions ---
    const file = DriveApp.getFileById(doc.getId());
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    recipients.forEach(r => { try { file.addViewer(r); } catch (err) {} });

    // --- 3. Log to Tracking Sheet ---
    try {
      if (LOGGING_SHEET_URL) {
        const logSheet = SpreadsheetApp.openByUrl(LOGGING_SHEET_URL).getSheets()[0];
        logSheet.appendRow([
          new Date(),           // Timestamp
          person,               // Strategist Name
          client,               // Client Name
          audience,             // Topic
          doc.getUrl(),         // Link to the Google Doc
          infoUrl,              // Link to the Infographic
          recipients.join(', ') // Who it was sent to
        ]);
      }
    } catch (logErr) {
      console.log("Logging failed, but continuing execution: " + logErr);
    }

    // --- 4. Send Professional, No-Reply Email ---
    if (recipients.length > 0) {
      const emailSubject = `Strategic Insights Report: ${audience} | Prepared for ${client}`;
      
      const emailBody = `Dear Team,\n\n` +
        `Please find the strategic AI insights report for ${audience}, prepared specifically for ${client}.\n\n` +
        `This document contains a comprehensive analysis of the latest search trends, competitive dynamics, and actionable marketing recommendations based on consumer intent.\n\n` +
        `📄 Access the Strategy Document here: ${doc.getUrl()}\n\n` +
        `Best regards,\n` +
        `${person}`;

      // Using MailApp instead of GmailApp to enable 'noReply'
      MailApp.sendEmail({
        to: recipients.join(','),
        subject: emailSubject,
        body: emailBody,
        name: "Insights Genie", // Display name in the inbox
        noReply: true           // Forces a noreply@ address
      });
    }

    return { ok: true, docUrl: doc.getUrl(), url: infoUrl };
  } catch (err) {
    return { ok: false, error: "Publishing Error: " + err.message };
  }
}

/***** ================= UTILS ================= *****/
function extractSheetId_(url) {
  const m = String(url || '').match(/[-\w]{25,}/);
  return m ? m[0] : null;
}

function setInfographicHTML(html) { return true; }
function setInfographicURL(url) { 
  PropertiesService.getUserProperties().setProperty(INFO_URL_PROP, url);
  return true; 
}

/***** ================= USER CONTEXT (Via Admin Directory) ================= *****/
function getCurrentUserContext() {
  const email = Session.getActiveUser().getEmail();
  let fullName = "";
  let photoUrl = "";

  try {
    const user = AdminDirectory.Users.get(email, { viewType: 'domain_public' });
    if (user && user.name) {
      fullName = user.name.fullName || user.name.givenName;
    }
  } catch (e) {
    console.log("Admin Directory lookup failed: " + e.message);
  }

  if (!fullName) {
    const handle = email.split('@')[0];
    fullName = handle.split('.').map(part => part.charAt(0).toUpperCase() + part.slice(1)).join(' ');
  }

  return { email: email, name: fullName, photoUrl: photoUrl };
}

/***** ================= FORCE AUTHORIZATION ================= *****/
function authorizeScript() {
  DriveApp.getRootFolder(); 
  GmailApp.getInboxThreads(0, 1);
  
  // NEW: Forces Google to detect the MailApp permission requirement
  if (false) { MailApp.sendEmail("test@google.com", "test", "test"); }

  const tempDoc = DocumentApp.create("Insights Genie Permission Test");
  tempDoc.saveAndClose();
  DriveApp.getFileById(tempDoc.getId()).setTrashed(true);
  try {
    const email = Session.getActiveUser().getEmail();
    AdminDirectory.Users.get(email, { viewType: 'domain_public' });
  } catch (e) {
    console.warn('Admin Directory check failed: ' + e.message);
  }
}

/***** ================= CORE: Generate Custom Miscellaneous Insights ================= *****/
function generateMiscellaneousInsights(sheetUrl, targetaudience, customPrompt) {
  try {
    const sheetId = extractSheetId_(sheetUrl);
    if (!sheetId) throw new Error('Invalid Sheet URL');

    const ss = SpreadsheetApp.openById(sheetId);
    let miscSheet = ss.getSheetByName("Miscellaneous") || ss.getSheetByName("Misllenious");
    if (!miscSheet) throw new Error('Could not find a tab named "Miscellaneous" in the provided sheet.');

    const data = miscSheet.getDataRange().getDisplayValues();
    if (data.length < 2) throw new Error('No data found in the Miscellaneous tab.');

    const csvString = data.map(row => row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')).join('\n');

    const prompt = `ACT AS A SENIOR DATA ANALYST AND MARKET RESEARCHER.\nTOPIC: ${targetaudience}\nINSTRUCTIONS: ${customPrompt}\nDATA:\n${csvString}`;
    return { ok: true, reportText: callGeminiText_(prompt) };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function testMail() {
  MailApp.sendEmail(Session.getActiveUser().getEmail(), "Auth Test", "This is just to force the permission popup.");
}/***** ================= Insights Geine – Full Restore + Resilience Logic ================= *****/

function doGet(e) {
  if (e && e.parameter && e.parameter.view) {
    try {
      const fileId = e.parameter.view;
      const file = DriveApp.getFileById(fileId);
      const htmlContent = file.getBlob().getDataAsString();

      return HtmlService.createHtmlOutput(htmlContent)
        .setTitle('Insights Geine Infographic')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (err) {
      return HtmlService.createHtmlOutput(
        "<h3>Error: Infographic not found or access denied.</h3>" +
        "<p>Please ensure the file exists and you have authorized Drive access.</p>"
      );
    }
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Insights Geine')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/***** ================= CONFIG ================= *****/
const GEMINI_KEY_PROP = 'GEMINI_API_KEY';
const GEMINI_MODEL = 'gemini-2.5-pro';
const GEMINI_BASE = 'https://generativelanguage.googleapis.com/v1beta/models/';
const DRIVE_FOLDER_ID = '1bB_2w_Dv36bIN3Kxr7WFNOcocicSz5A9';
const INFO_URL_PROP = 'Insights_Geine_LAST_INFO_URL';

/***** ================= CORE: Generate AI Insights ================= *****/
function generateInsightsFromSheet(sheetUrl, targetaudience) {
 try {
    const sheetId = extractSheetId_(sheetUrl);
    if (!sheetId) throw new Error('Invalid Sheet URL');

    const ss = SpreadsheetApp.openById(sheetId);
    
    // --- 1. PROCESS TERMS DATA (First Tab) ---
    const termSheet = ss.getSheets()[0];
    const termData = termSheet.getDataRange().getValues();
    if (termData.length < 2) throw new Error('No terms data found.');

    const tHeaders = termData[0].map(h => h.toString().toLowerCase());
    const qI = tHeaders.findIndex(h => h.includes('query'));
    const oI = tHeaders.findIndex(h => h.includes('opportunities') && !h.includes('pop'));
    const pI = tHeaders.findIndex(h => h.includes('pop'));

    if (qI === -1 || oI === -1 || pI === -1) throw new Error('Missing headers in Terms sheet: query, opportunity, PoP');

    const rows = termData.slice(1).map(r => ({
      query: String(r[qI] || '').trim(),
      adOpp: parseFloat(r[oI]) || 0,
      // FIX: Multiply by 100 because Sheets returns percentages as decimals (e.g., 0.5 for 50%)
      adOppPoP: (parseFloat(r[pI]) || 0) * 100 
    })).filter(r => r.query);

    const maxOpp = Math.max(...rows.map(r => r.adOpp));
    
    // Create the TERMS bundle
    const termEnriched = rows.map(r => ({
      query: r.query,
      IOS: maxOpp ? (r.adOpp / maxOpp) * 100 : 0,
      PoP: r.adOppPoP
    })).sort((a,b) => b.IOS - a.IOS);

    const termBundle = { top: termEnriched.slice(0, 100), totalRows: termEnriched.length };

    // --- 2. PROCESS LOCATION DATA (Tab: "Location_Data") ---
    let locationBundle = [];
    const locSheet = ss.getSheetByName("Location_Data");
    
    if (locSheet) {
      const locData = locSheet.getDataRange().getValues();
      if (locData.length > 1) {
        const lHeaders = locData[0].map(h => h.toString().toLowerCase());
        const gI = lHeaders.findIndex(h => h.includes('geo') || h.includes('name'));
        const loI = lHeaders.findIndex(h => h.includes('opportunities') && !h.includes('pop'));
        const lpI = lHeaders.findIndex(h => h.includes('pop'));

        if (gI > -1 && loI > -1 && lpI > -1) {
          locationBundle = locData.slice(1).map(r => ({
            geo: String(r[gI]),
            volume: parseFloat(r[loI]) || 0,
            // FIX: Multiply Location Growth by 100 as well
            growth: (parseFloat(r[lpI]) || 0) * 100
          }));
        }
      }
    }


    // --- 3. PROCESS INTENT DATA (Tab: "Intent" or "Intent_Data") ---
    let intentBundle = [];
    const intentSheet = ss.getSheetByName("Intent") || ss.getSheetByName("Intent_Data");
    if (intentSheet) {
      const intentData = intentSheet.getDataRange().getValues();
      if (intentData.length > 1) {
        const iHeaders = intentData[0].map(h => h.toString().toLowerCase());
        const iName = iHeaders.findIndex(h => h.includes('intent') || h.includes('query') || h.includes('name'));
        const iOpp = iHeaders.findIndex(h => h.includes('opportunities') && !h.includes('pop'));
        const iPop = iHeaders.findIndex(h => h.includes('pop'));
        
        intentBundle = intentData.slice(1).map(r => ({
          intent: String(r[iName > -1 ? iName : 0]),
          volume: parseFloat(r[iOpp > -1 ? iOpp : 1]) || 0,
          growth: (parseFloat(r[iPop > -1 ? iPop : 2]) || 0) * 100
        })).filter(r => r.intent);
      }
    }

    // --- 4. PROCESS FEATURES DATA (Tab: "Features" or "Features_Data") ---
    let featureBundle = [];
    const featureSheet = ss.getSheetByName("Features") || ss.getSheetByName("Features_Data");
    if (featureSheet) {
      const featureData = featureSheet.getDataRange().getValues();
      if (featureData.length > 1) {
        const fHeaders = featureData[0].map(h => h.toString().toLowerCase());
        const fName = fHeaders.findIndex(h => h.includes('feature') || h.includes('query') || h.includes('name'));
        const fOpp = fHeaders.findIndex(h => h.includes('opportunities') && !h.includes('pop'));
        const fPop = fHeaders.findIndex(h => h.includes('pop'));
        
        featureBundle = featureData.slice(1).map(r => ({
          feature: String(r[fName > -1 ? fName : 0]),
          volume: parseFloat(r[fOpp > -1 ? fOpp : 1]) || 0,
          growth: (parseFloat(r[fPop > -1 ? fPop : 2]) || 0) * 100
        })).filter(r => r.feature);
      }
    }

    // --- 5. PROCESS BRANDS DATA (Tab: "Brands" or "Brands_Data") ---
    let brandBundle = [];
    const brandSheet = ss.getSheetByName("Brands") || ss.getSheetByName("Brands_Data");
    if (brandSheet) {
      const brandData = brandSheet.getDataRange().getValues();
      if (brandData.length > 1) {
        const bHeaders = brandData[0].map(h => h.toString().toLowerCase());
        const bName = bHeaders.findIndex(h => h.includes('brand') || h.includes('query') || h.includes('name'));
        const bOpp = bHeaders.findIndex(h => h.includes('opportunities') && !h.includes('pop'));
        const bPop = bHeaders.findIndex(h => h.includes('pop'));
        
        brandBundle = brandData.slice(1).map(r => ({
          brand: String(r[bName > -1 ? bName : 0]),
          volume: parseFloat(r[bOpp > -1 ? bOpp : 1]) || 0,
          growth: (parseFloat(r[bPop > -1 ? bPop : 2]) || 0) * 100
        })).filter(r => r.brand);
      }
    }


    const prompt = `
ACT AS A SENIOR MARKET RESEARCH LEAD AND GOOGLE STRATEGIST. Create a high-impact, structured insight report based on the provided search data.
TOPIC: ${targetaudience}

DATA DEFINITIONS:
Topic: The keyword, feature, intent, or brand used to find this product.
Searches (Ad Opps): The raw volume of interest. CRITICAL FILTER: IGNORE any query with fewer than 1,000 Searches. Do not include these in any analysis.
YoY Growth: The percentage change in interest over the last year.

RAW DATASET (General Terms):
(Top 100 of ${termBundle.totalRows} queries):
${JSON.stringify(termBundle.top, null, 2)}

LOCATION DATA (Regional Breakdown):
${JSON.stringify(locationBundle, null, 2)}

INTENT DATA:
${JSON.stringify(intentBundle, null, 2)}

FEATURES DATA:
${JSON.stringify(featureBundle, null, 2)}

BRANDS DATA:
${JSON.stringify(brandBundle, null, 2)}

REPORT STRUCTURE:

1. TOP 3 EMERGING MARKET TRENDS
Identify overarching themes in the data (using only terms with >1,000 searches). For each trend, provide:
Trend Name: A descriptive title for the shift.
Metrics: Use [YoY Growth %] to justify the trend.
Evidence: Cite specific queries from the data that prove this trend exists for ${targetaudience}.

2. Brand Share of Searches
Using the search data, identify the presence of specific brands versus generic terms.
Relative Share Calculation: Do NOT use raw volumes. Instead, calculate the "Relative Share" percentage within the competitive set found in the data (e.g., if Brand A, B, and C are found, Brand A share = Brand A Volume / Total Volume of Brands A+B+C).
Analysis: Highlight which brands are dominating the Share of Mind for ${targetaudience} and which are losing ground based on YoY growth.
Format/Output: Present a clean breakdown showing the Calculated Share of Search (%) and the YoY Growth (%) for each brand. Add a brief qualitative summary of the competitive landscape (e.g., Who is the undisputed leader? Who are the rising challengers?).

3. BRAND X ATTRIBUTE
STRICT DATA ISOLATION RULE: For this section, you are strictly forbidden from looking at the "Brands", "Intent", or "Features" data tabs. You MUST derive the Brand x Attribute resonance EXCLUSIVELY by analyzing the granular queries within the "Top terms" data.
Objective: Find granular queries in the top terms where a specific brand is searched alongside a specific feature or attribute (e.g., a high-volume query like "[Brand A] camera" vs "[Brand B] battery"). Provide strong qualitative commentary on which brands are successfully owning specific narratives or attributes in the minds of consumers.
Format: Identify 3-4 interesting intersections directly from the query strings. Explain the consumer perception based on the search volume and growth of those specific raw queries.

4. The Intent Landscape
Analyze the "Intent" data to map where consumers are in the funnel.
Objective: Calculate the Share of Searches for different consumer intents relative to the total intent volume using Ad Opportunities.
Format: Present the Intent Share (%) and YoY Growth (%). Add commentary on shifting funnel dynamics (e.g., Is there a massive rise in "Comparison" or "Refurbished" queries indicating price sensitivity? Are "Review" searches outpacing "Best" searches?).

5.  Feature Demand Breakdown
Analyze the "Features" data to show what hardware/software capabilities matter most.
Objective: Calculate the Share of Searches for specific smartphone features relative to the total feature volume using Ad Opportunities.Have only top 6 by Ad Opportunities volume - the rest you can put in others
Format: Present the Feature Share (%) and YoY Growth (%). Add qualitative insights on which features are now considered "table stakes" (high share, flat growth) versus "innovation drivers" (lower share, explosive growth).

6. Strategic Marketing Recommendations (Google Ads & Content)
Translate the above insights into actionable sales pitches for the client.
Objective: Provide 4-5 concrete marketing recommendations.
Format: Group recommendations by strategy.
Content Strategy: How should brands pivot their ad creatives, video narratives, and messaging pillars based on the Features and Intent data?
Google Ads Strategy: How should they structure their Google Ads campaigns? (e.g., Bidding aggressively on mid-funnel comparison terms via Search, using explosive feature trends as hooks in YouTube Shorts campaigns, or leveraging Performance Max to capture shifting brand allegiances).

IMPORTANT:
Do NOT start with meta sentences like "Of course..." or "Here is your report".
Start directly from section 1.
No conversational filler.
STRICTLY PLAIN TEXT: Do NOT use Markdown formatting (no asterisks like ** or *).
Do NOT use bolding syntax.
Use uppercase for main headings (e.g., "1. TOP 3 EMERGING MARKET TRENDS") to distinguish them without bolding symbols.
THRESHOLD RULE: STRICTLY ignore all data points with "Searches" < 1,000.
VOLUME RULE: Never output raw search volume numbers. Use "Relative Share %" for groups/competitors and "YoY Growth %" elsewhere.
Ensure the advice is hyper-specific to ${targetaudience}.
`;

    const text = callGeminiText_(prompt);
    return { ok: true, reportText: text };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

/***** ================= GEMINI TEXT ================= *****/
function callGeminiText_(prompt) {
  const key = PropertiesService.getScriptProperties().getProperty(GEMINI_KEY_PROP);
  const url = `${GEMINI_BASE}${GEMINI_MODEL}:generateContent?key=${key}`;
  const payload = {
    contents: [{ role: 'user', parts: [{ text: prompt }]}],
    generationConfig: { temperature: 0.3, maxOutputTokens: 50000 }
  };
  const resp = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true, timeout: 60000
  });
  const parsed = JSON.parse(resp.getContentText() || '{}');
  return parsed.candidates?.[0]?.content?.parts?.[0]?.text || '';
}

/***** ================= GENERATE INFOGRAPHIC HTML ================= *****/
function generateInfographicHtmlFromInsights(targetaudience, reportText) {
  try {
    const key = PropertiesService.getScriptProperties().getProperty(GEMINI_KEY_PROP);
    const url = `${GEMINI_BASE}${GEMINI_MODEL}:generateContent?key=${key}`;

    const prompt = `You are generating a polished, modern Tailwind infographic.

STRICT OUTPUT RULES:
- Output ONLY HTML (no markdown, no code fences).
- Do NOT include <html>, <head>, <body>, or <script>.
- Use Tailwind utility classes only.
- Keep text concise. No long paragraphs.
- Like in the below example always keep the footer in the end.
- Also add the some explanation regarding the metric 
- Make sure that the infographics should contain all the insights while generating it.
- Always add the variety of chart wherever it make sense for each sections for eg table chart,pie hart,histogram ,bar chart etc.
- Horizontal Scroller: Applied to the "Emerging Market Trends" (Section 1) so users can swipe/scroll left to right through the trends.
- Horizontal Bar Chart: Applied to the "Brand Share" (Section 2) for a clean competitive landscape view.
- Brand Resonance Matrix: Enhanced the "Brand x Attribute" (Section 3) into a more distinct grid format.
- Donut Chart: Applied to "The Intent Landscape" (Section 4).
- Radar Chart: Applied to the "Feature Demand Breakdown" (Section 5) to visualize how different hardware features stack up against each other.
- Action Cards: Upgraded the "Strategic Marketing Recommendations" (Section 6) to look like distinct, actionable cards.
- "We recommend cross-checking all insights before sending to clients." this line should always be at the footer of the infographics

Maintain the theme of the example below across all infographics, but use complete insights:
${reportText} which is generated to make infographics:

<!DOCTYPE html><html><head><script src="https://cdn.tailwindcss.com"></script></head><body style="background:#f8fafc; padding:24px;"><div style="max-width:1152px; margin:0 auto; background:white; border-radius:16px; padding:32px; box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1);"><div class="max-w-6xl mx-auto bg-white border border-slate-200 rounded-2xl shadow-sm p-6 md:p-8 font-sans text-slate-800">

    <header class="max-w-6xl mx-auto mb-10 text-center">
        <div class="inline-block bg-blue-600 text-white px-4 py-1 rounded-full text-sm font-semibold tracking-wide uppercase mb-3">Search Insights Report</div>
        <h1 class="text-4xl md:text-5xl font-extrabold text-slate-900 mb-4">The Smartphone Market</h1>
        <p class="text-lg text-slate-600 max-w-3xl mx-auto">An analysis of consumer search trends, brand dynamics, and feature demand shaping the future of mobile technology.</p>
    </header>

    <main class="max-w-6xl mx-auto space-y-12">

        <!-- Section 1: Emerging Market Trends -->
        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">1. Top 3 Emerging Market Trends</h2>
                    <p class="text-slate-400 mt-1">Consumer behavior is radically shifting towards future-planning and practical AI application.</p>
                </div>
                <svg class="w-10 h-10 text-blue-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 7h8m0 0v8m0-8l-8 8-4-4-6 6"></path></svg>
            </div>
            
            <div class="relative">
                <div class="flex overflow-x-auto snap-x snap-mandatory scrollbar-thin scrollbar-thumb-slate-300 scrollbar-track-slate-100 p-6 md:p-8 space-x-6">
                    
                    <!-- Trend 1 Card -->
                    <div class="snap-center flex-shrink-0 w-11/12 md:w-2/3 lg:w-1/2 bg-slate-50 p-6 rounded-xl border border-slate-200">
                        <h3 class="text-lg font-semibold text-slate-800 mb-2">Pre-Launch Hype Cycle Acceleration</h3>
                        <p class="text-slate-600 mb-4 text-sm">Consumers are now planning purchases years in advance. Searches for unreleased models show massive YoY growth, indicating a need to engage audiences long before a product launch.</p>
                        <div class="grid grid-cols-2 gap-4">
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"iphone 17"</div><div class="text-2xl font-bold text-green-600 mt-1">+1398%</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"oppo reno 15"</div><div class="text-2xl font-bold text-green-600 mt-1">+109k%</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"iphone 17 pro max"</div><div class="text-2xl font-bold text-green-600 mt-1">+1023%</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"vivo v60"</div><div class="text-2xl font-bold text-green-600 mt-1">+27k%</div></div>
                        </div>
                    </div>

                    <!-- Trend 2 Card -->
                    <div class="snap-center flex-shrink-0 w-11/12 md:w-2/3 lg:w-1/2 bg-slate-50 p-6 rounded-xl border border-slate-200">
                        <h3 class="text-lg font-semibold text-slate-800 mb-2">The AI Utility Shift</h3>
                        <p class="text-slate-600 mb-4 text-sm">The conversation has moved from "What is AI?" to "How do I use it?". General AI queries are declining, while specific, use-case-oriented searches are exploding.</p>
                        <div class="grid grid-cols-1 gap-4">
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"new galaxy ai phone"</div><div class="text-2xl font-bold text-green-600 mt-1">+17,732% YoY</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"how to use circle to search"</div><div class="text-2xl font-bold text-green-600 mt-1">+821% YoY</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"writing assist samsung"</div><div class="text-2xl font-bold text-green-600 mt-1">+6000% YoY</div></div>
                        </div>
                    </div>

                    <!-- Trend 3 Card -->
                    <div class="snap-center flex-shrink-0 w-11/12 md:w-2/3 lg:w-1/2 bg-slate-50 p-6 rounded-xl border border-slate-200">
                        <h3 class="text-lg font-semibold text-slate-800 mb-2">Demand for Hyper-Specific Innovation</h3>
                        <p class="text-slate-600 mb-4 text-sm">Users are now researching niche, technical features that have a tangible impact on daily use, signaling a more sophisticated and demanding consumer base.</p>
                        <div class="grid grid-cols-1 gap-4">
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"privacy display samsung"</div><div class="text-2xl font-bold text-green-600 mt-1">+314,600% YoY</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"what is gemini on my phone"</div><div class="text-2xl font-bold text-green-600 mt-1">+46,100% YoY</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"priority notification iphone"</div><div class="text-2xl font-bold text-green-600 mt-1">+1558% YoY</div></div>
                        </div>
                    </div>
                </div>
            </div>
        </section>

        <!-- Section 2: Brand Share -->
        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">2. Brand Share of Searches</h2>
                    <p class="text-slate-400 mt-1">A two-tiered market where leaders solidify their position while challengers show dynamic growth.</p>
                </div>
                <svg class="w-10 h-10 text-indigo-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"></path></svg>
            </div>
            
            <div class="p-6 md:p-8 grid grid-cols-1 lg:grid-cols-5 gap-8">
                <div class="lg:col-span-3">
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Calculated Share & YoY Growth</h3>
                    <div class="space-y-4 text-sm">
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Apple</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 27.31%;">27.3%</div></div><div class="w-20 text-right font-medium text-green-600">+13.34%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Samsung</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 18.23%;">18.2%</div></div><div class="w-20 text-right font-medium text-red-600">-4.08%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">vivo</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 13.46%;">13.5%</div></div><div class="w-20 text-right font-medium text-green-600">+18.60%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">OPPO</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 8.76%;">8.8%</div></div><div class="w-20 text-right font-medium text-green-600">+20.53%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Realme</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 6.99%;">7.0%</div></div><div class="w-20 text-right font-medium text-red-600">-2.63%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Redmi</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 4.87%;">4.9%</div></div><div class="w-20 text-right font-medium text-red-600">-25.52%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Motorola</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 4.27%;">4.3%</div></div><div class="w-20 text-right font-medium text-green-600">+12.89%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Google</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 2.34%;">2.3%</div></div><div class="w-20 text-right font-medium text-green-600">+27.61%</div></div>
                    </div>
                </div>
                <div class="lg:col-span-2">
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Market Summary</h3>
                    <div class="space-y-3 text-sm text-slate-600">
                        <p><span class="font-bold text-slate-800">Leaders:</span> Apple solidifies its #1 spot with strong growth. Samsung maintains a huge share but shows signs of erosion.</p>
                        <p><span class="font-bold text-slate-800">Challengers:</span> Vivo and OPPO are standout growers, bucking the negative trend seen by competitors like Redmi and OnePlus.</p>
                        <p><span class="font-bold text-slate-800">Movers:</span> Google's Pixel line is gaining traction with notable growth. Huawei is staging a powerful comeback (+79% YoY), signaling a resurgence in consumer interest.</p>
                    </div>
                </div>
            </div>
        </section>

        <!-- Section 3: Brand x Attribute -->
        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">3. Brand x Attribute Resonance</h2>
                    <p class="text-slate-400 mt-1">How consumers perceive and search for brand-specific innovations.</p>
                </div>
                <svg class="w-10 h-10 text-teal-400" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M9.568 3H5.25A2.25 2.25 0 003 5.25v4.318c0 .597.237 1.17.659 1.591l9.581 9.581c.699.699 1.78.872 2.607.33a18.095 18.095 0 005.223-5.223c.542-.827.369-1.908-.33-2.607L11.16 3.66A2.25 2.25 0 009.568 3z" /><path stroke-linecap="round" stroke-linejoin="round" d="M6 6h.008v.008H6V6z" /></svg>
            </div>
            
            <div class="p-6 md:p-8 grid grid-cols-1 lg:grid-cols-3 gap-6">
                <div class="bg-slate-50 p-6 rounded-xl border border-slate-200">
                    <h3 class="text-lg font-semibold text-slate-800 mb-2">Samsung Owns "Next-Gen AI Utility"</h3>
                    <p class="text-slate-600 text-sm">Samsung has successfully branded its AI suite. Consumers associate them with tangible features, searching for <span class="font-semibold text-teal-700">"circle to search samsung"</span> and <span class="font-semibold text-teal-700">"samsung ai phone"</span>, proving they own the narrative of practical, in-hand innovation.</p>
                </div>
                <div class="bg-slate-50 p-6 rounded-xl border border-slate-200">
                    <h3 class="text-lg font-semibold text-slate-800 mb-2">Apple's Narrative: "Future Hype & AI Compatibility"</h3>
                    <p class="text-slate-600 text-sm">Apple dominates future product hype (<span class="font-semibold text-teal-700">"iphone 17"</span>). Their AI narrative is focused on compatibility, with queries like <span class="font-semibold text-teal-700">"apple intelligence supported phones"</span>, indicating a user mindset focused on access rather than specific use-cases.</p>
                </div>
                <div class="bg-slate-50 p-6 rounded-xl border border-slate-200">
                    <h3 class="text-lg font-semibold text-slate-800 mb-2">The Rise of the Hyper-Informed Searcher</h3>
                    <p class="text-slate-600 text-sm">Users are making decisions before they search. Highly detailed, near-purchase queries like <span class="font-semibold text-teal-700">"samsung galaxy s24 fe 5g ai smartphone"</span> show that brands must match this specificity to capture high-intent customers.</p>
                </div>
            </div>
        </section>

        <!-- Section 4: Intent Landscape -->
        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">4. The Intent Landscape</h2>
                    <p class="text-slate-400 mt-1">Price sensitivity is paramount, with explosive growth in value-based acquisition models.</p>
                </div>
                <svg class="w-10 h-10 text-amber-400" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M21 21l-5.197-5.197m0 0A7.5 7.5 0 105.196 5.196a7.5 7.5 0 0010.607 10.607z" /></svg>
            </div>
            
            <div class="p-6 md:p-8 grid grid-cols-1 lg:grid-cols-2 gap-8 items-center">
                <div class="relative w-64 h-64 mx-auto">
                    <div class="absolute inset-0 rounded-full" style="background-image: conic-gradient(from 0deg, #f59e0b 0% 50.43%, #64748b 50.43% 64.07%, #fbbf24 64.07% 77.42%, #fcd34d 77.42% 86.43%, #fef3c7 86.43% 100%);"></div>
                    <div class="absolute inset-5 bg-white rounded-full flex flex-col items-center justify-center">
                        <div class="text-4xl font-bold text-slate-800">50.4%</div>
                        <div class="text-sm text-slate-500 font-semibold">Price Intent</div>
                    </div>
                </div>
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Funnel Dynamics</h3>
                    <p class="text-slate-600 text-sm mb-4">While "Price" searches dominate, the real story is the massive growth in action-oriented, value-seeking behaviors. The decline in "Best" and "Reviews" suggests consumers are tired of generic advice and are focusing on the financial aspect of their purchase.</p>
                    <div class="grid grid-cols-2 gap-x-4 gap-y-2 text-sm">
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-amber-500 mr-2"></span>Price: <span class="ml-auto font-medium text-green-600">+4.2%</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-slate-500 mr-2"></span>Best: <span class="ml-auto font-medium text-red-600">-11.9%</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-amber-400 mr-2"></span>Comparison: <span class="ml-auto font-medium text-green-600">+4.1%</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-slate-500 mr-2"></span>Reviews: <span class="ml-auto font-medium text-red-600">-5.6%</span></div>
                        <div class="flex items-center col-span-2 border-t pt-2 mt-2 border-slate-200"><span class="font-bold text-slate-800">Growth Drivers:</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-green-500 mr-2"></span>Deals: <span class="ml-auto font-bold text-green-600">+93%</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-green-500 mr-2"></span>Used: <span class="ml-auto font-bold text-green-600">+123%</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-green-500 mr-2"></span>Trade-In: <span class="ml-auto font-bold text-green-600">+141%</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-green-500 mr-2"></span>Contract: <span class="ml-auto font-bold text-green-600">+187%</span></div>
                    </div>
                </div>
            </div>
        </section>

        <!-- Section 5: Feature Demand -->
        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">5. Feature Demand Breakdown</h2>
                    <p class="text-slate-400 mt-1">5G is now a baseline expectation. The new battlegrounds are screen technology and productivity.</p>
                </div>
                <svg class="w-10 h-10 text-rose-400" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M10.5 6h9.75M10.5 6a1.5 1.5 0 11-3 0m3 0a1.5 1.5 0 10-3 0M3.75 6H7.5m3 12h9.75m-9.75 0a1.5 1.5 0 01-3 0m3 0a1.5 1.5 0 00-3 0m-3.75 0H7.5m9-6h3.75m-3.75 0a1.5 1.5 0 01-3 0m3 0a1.5 1.5 0 00-3 0m-9.75 0h9.75" /></svg>
            </div>
            
            <div class="p-6 md:p-8 grid grid-cols-1 lg:grid-cols-2 gap-8 items-center">
                <div class="relative w-80 h-80 mx-auto">
                    <svg viewBox="0 0 200 200" class="w-full h-full">
                        <!-- Grid Lines -->
                        <polygon points="100,20 175,55 175,145 100,180 25,145 25,55" fill="none" stroke="#e2e8f0" stroke-width="1"/>
                        <polygon points="100,40 160,67 160,133 100,160 40,133 40,67" fill="none" stroke="#e2e8f0" stroke-width="1"/>
                        <polygon points="100,60 145,80 145,120 100,140 55,120 55,80" fill="none" stroke="#e2e8f0" stroke-width="1"/>
                        <line x1="100" y1="20" x2="100" y2="180" stroke="#e2e8f0" stroke-width="1"/>
                        <line x1="25" y1="55" x2="175" y2="145" stroke="#e2e8f0" stroke-width="1"/>
                        <line x1="25" y1="145" x2="175" y2="55" stroke="#e2e8f0" stroke-width="1"/>
                        <!-- Data Polygon -->
                        <polygon points="100,104.14 116.3,130.65 107.99,142.01 100,132.53 86.66,133.34 85.82,113.34" fill="rgba(244, 63, 94, 0.2)" stroke="#f43f5e" stroke-width="2"/>
                        <!-- Labels -->
                        <text x="100" y="15" text-anchor="middle" font-size="10" fill="#475569" font-weight="600">5G (-18%)</text>
                        <text x="180" y="55" text-anchor="middle" font-size="10" fill="#475569" font-weight="600">Screen (+102%)</text>
                        <text x="180" y="150" text-anchor="middle" font-size="10" fill="#475569" font-weight="600">Camera (+9%)</text>
                        <text x="100" y="195" text-anchor="middle" font-size="10" fill="#475569" font-weight="600">Color (+58%)</text>
                        <text x="20" y="150" text-anchor="middle" font-size="10" fill="#475569" font-weight="600">Battery (+21%)</text>
                        <text x="20" y="55" text-anchor="middle" font-size="10" fill="#475569" font-weight="600">Storage (+3%)</text>
                    </svg>
                </div>
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Feature Insights</h3>
                    <div class="space-y-4 text-sm text-slate-600">
                        <div>
                            <div class="font-bold text-slate-800">Table Stakes</div>
                            <p>"5G" has high search share but is declining, meaning consumers now expect it as standard. "Camera," "Battery," and "Storage" are evergreen pillars that require continuous competitive messaging.</p>
                        </div>
                        <div>
                            <div class="font-bold text-slate-800">Innovation Drivers</div>
                            <p>"Screen" is the breakout star (+102% YoY), showing that advancements in refresh rates, brightness, and privacy are capturing significant interest. "Color" (+58% YoY) also remains a powerful and growing differentiator.</p>
                        </div>
                    </div>
                </div>
            </div>
        </section>

        <!-- Section 6: Recommendations -->
        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">6. Strategic Marketing Recommendations</h2>
                    <p class="text-slate-400 mt-1">Actionable strategies for Google Ads and Content to capitalize on market shifts.</p>
                </div>
                <svg class="w-10 h-10 text-cyan-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9.663 17h4.673M12 3v1m6.364 1.636l-.707.707M21 12h-1M4 12H3m3.343-5.657l-.707-.707m2.828 9.9a5 5 0 117.072 0l-.548.547A3.374 3.374 0 0014 18.469V19a2 2 0 11-4 0v-.531c0-.895-.356-1.754-.988-2.386l-.548-.547z"></path></svg>
            </div>
            
            <div class="p-6 md:p-8 grid grid-cols-1 md:grid-cols-2 gap-8">
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Content Strategy</h3>
                    <div class="space-y-4">
                        <div class="p-4 bg-slate-50 rounded-lg border border-slate-200">
                            <div class="font-bold text-slate-800">1. Pivot AI Messaging from "What" to "How"</div>
                            <p class="text-sm text-slate-600 mt-1">Create content focused on utility. Develop guides and Shorts like "5 Ways Galaxy AI Can Organize Your Week" to capture the high-growth "how-to" search trend.</p>
                        </div>
                        <div class="p-4 bg-slate-50 rounded-lg border border-slate-200">
                            <div class="font-bold text-slate-800">2. Build Content Hubs for Innovation Drivers</div>
                            <p class="text-sm text-slate-600 mt-1">Develop deep content around "Screen Technology." Go beyond specs with articles on "How Privacy Screens Work" to capture sophisticated users searching for high-growth features.</p>
                        </div>
                    </div>
                </div>
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Google Ads Strategy</h3>
                    <div class="space-y-4">
                        <div class="p-4 bg-slate-50 rounded-lg border border-slate-200">
                            <div class="font-bold text-slate-800">3. Restructure Campaigns Around New Acquisition Models</div>
                            <p class="text-sm text-slate-600 mt-1">Build dedicated campaigns for "Contract," "Trade-In," and "Used" keywords. Lead ad copy with value: "Upgrade with a $400 Trade-In Bonus" to address market price sensitivity.</p>
                        </div>
                        <div class="p-4 bg-slate-50 rounded-lg border border-slate-200">
                            <div class="font-bold text-slate-800">4. Capture Switchers with Feature-Led PMax</div>
                            <p class="text-sm text-slate-600 mt-1">Use Performance Max to target "Comparison" searches. Use competitor brand names as audience signals and showcase high-growth features like vibrant "Colors" and "Screen" quality.</p>
                        </div>
                        <div class="p-4 bg-slate-50 rounded-lg border border-slate-200">
                            <div class="font-bold text-slate-800">5. Build Future Audiences with Pre-Launch Campaigns</div>
                            <p class="text-sm text-slate-600 mt-1">Capitalize on searches for "iPhone 17." Run light-touch YouTube/Discovery campaigns to build remarketing lists, then retarget them with pre-order messaging closer to launch.</p>
                        </div>
                    </div>
                </div>
            </div>
        </section>

    </main>

    <footer class="text-center text-sm text-slate-500 mt-12 pt-8 border-t border-slate-200">
        <p class="mt-2 text-slate-400 font-medium">We recommend cross-checking all insights before sending to clients.</p>
    </footer>

</div></div></body></html>

TITLE: ${targetaudience}

`;
  const payload = {
      contents: [{ role: 'user', parts: [{ text: prompt }]}],
      // Lowered maxOutputTokens slightly to prevent API truncation errors
      generationConfig: { temperature: 0.2, maxOutputTokens: 30000 }
    };

    const resp = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload), muteHttpExceptions: true, timeout: 60000
    });

    const responseText = resp.getContentText();
    let parsed = {};
    
    // Safely parse JSON in case the API throws an HTML error page (like a 502)
    try {
      parsed = JSON.parse(responseText);
    } catch (e) {
      throw new Error("Invalid API response format. The Gemini service might be overloaded.");
    }

    // Check for explicit API errors
    if (parsed.error) {
      throw new Error(`Gemini API Error: ${parsed.error.message}`);
    }

    // Extract content safely
    let htmlContent = parsed?.candidates?.[0]?.content?.parts?.[0]?.text || '';
    
    if (!htmlContent) {
      // Check if it was blocked by safety settings
      if (parsed?.candidates?.[0]?.finishReason === "SAFETY") {
        throw new Error("Content blocked by Google Safety Filters.");
      }
      throw new Error("Gemini returned an empty response. Try again.");
    }

    htmlContent = htmlContent.replace(/```html/gi, '').replace(/```/g, '').trim();

    // Pass to the Drive saver function
    const hostingResult = saveHtmlAndGetUrl_(targetaudience, htmlContent);
    return { ok: true, html: htmlContent, url: hostingResult.url };
    
  } catch (err) {
    // Send the EXACT error back to the frontend
    return { ok: false, error: err.message };
  }
}

/***** ================= HELPER: HOST ON DRIVE (REPAIRED) ================= *****/
function saveHtmlAndGetUrl_(label, htmlBody) {
  const fullHtml = `<!DOCTYPE html><html><head><script src="https://cdn.tailwindcss.com"></script></head><body style="background:#f8fafc; padding:24px;"><div style="max-width:1152px; margin:0 auto; background:white; border-radius:16px; padding:32px; box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1);">${htmlBody}</div></body></html>`;
  
  let folder;
  try {
    folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  } catch (e) {
    folder = DriveApp.getRootFolder();
  }
  
  const fileName = `Infographic-${label}-${Date.now()}.html`;
  const file = folder.createFile(fileName, fullHtml, MimeType.HTML);
  
  try {
    // Attempt 1: Domain-wide sharing (Best for corporate Workspaces)
    file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e) {
    try {
      // Attempt 2: Public Link fallback
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e2) {
      console.warn("Could not set sharing permissions. File remains private to the creator.");
    }
  }

  const scriptUrl = ScriptApp.getService().getUrl();
  // Ensure we format the URL correctly
  const finalUrl = scriptUrl ? `${scriptUrl}?view=${file.getId()}` : file.getUrl();

  return { url: finalUrl, fileId: file.getId() };
}

/***** ================= PUBLISH & EMAIL LOGIC ================= *****/
const LOGGING_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1ZfsNLTHFNitPQLxU1s0PHlR3y6z_DPtluldcZTEjyeU/edit';

function publishReportAndLog(audience, reportText, emailsCsv, person, client, rev) {
  try {
    const recipients = (emailsCsv || '').split(',').map(x => x.trim()).filter(Boolean);
    const userProps = PropertiesService.getUserProperties();
    const infoUrl = userProps.getProperty(INFO_URL_PROP) || '';

    // --- 1. Create Doc ---
    const doc = DocumentApp.create(`Insights Genie - ${audience} - ${new Date().toDateString()}`);
    const body = doc.getBody();
    body.appendParagraph(`AI Strategy Insights — ${audience}`).setHeading(DocumentApp.ParagraphHeading.TITLE);
    body.appendParagraph(`Strategist: ${person} | Client: ${client}`).setHeading(DocumentApp.ParagraphHeading.SUBTITLE);
    body.appendParagraph(reportText || '');

    if (infoUrl && infoUrl.startsWith('http')) {
      body.appendParagraph('Infographic').setHeading(DocumentApp.ParagraphHeading.HEADING2);
      const p = body.appendParagraph('📊 Open Infographic (Interactive)');
      p.setLinkUrl(infoUrl);
      p.setBold(true);
      p.setFontSize(13);
      p.setForegroundColor('#1a73e8');
    }
    doc.saveAndClose();

    // --- 2. Set Permissions ---
    const file = DriveApp.getFileById(doc.getId());
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    recipients.forEach(r => { try { file.addViewer(r); } catch (err) {} });

    // --- 3. Log to Tracking Sheet ---
    try {
      if (LOGGING_SHEET_URL) {
        const logSheet = SpreadsheetApp.openByUrl(LOGGING_SHEET_URL).getSheets()[0];
        logSheet.appendRow([
          new Date(),           // Timestamp
          person,               // Strategist Name
          client,               // Client Name
          audience,             // Topic
          doc.getUrl(),         // Link to the Google Doc
          infoUrl,              // Link to the Infographic
          recipients.join(', ') // Who it was sent to
        ]);
      }
    } catch (logErr) {
      console.log("Logging failed, but continuing execution: " + logErr);
    }

    // --- 4. Send Professional, No-Reply Email ---
    if (recipients.length > 0) {
      const emailSubject = `Strategic Insights Report: ${audience} | Prepared for ${client}`;
      
      const emailBody = `Dear Team,\n\n` +
        `Please find the strategic AI insights report for ${audience}, prepared specifically for ${client}.\n\n` +
        `This document contains a comprehensive analysis of the latest search trends, competitive dynamics, and actionable marketing recommendations based on consumer intent.\n\n` +
        `📄 Access the Strategy Document here: ${doc.getUrl()}\n\n` +
        `Best regards,\n` +
        `${person}`;

      // Using MailApp instead of GmailApp to enable 'noReply'
      MailApp.sendEmail({
        to: recipients.join(','),
        subject: emailSubject,
        body: emailBody,
        name: "Insights Genie", // Display name in the inbox
        noReply: true           // Forces a noreply@ address
      });
    }

    return { ok: true, docUrl: doc.getUrl(), url: infoUrl };
  } catch (err) {
    return { ok: false, error: "Publishing Error: " + err.message };
  }
}

/***** ================= UTILS ================= *****/
function extractSheetId_(url) {
  const m = String(url || '').match(/[-\w]{25,}/);
  return m ? m[0] : null;
}

function setInfographicHTML(html) { return true; }
function setInfographicURL(url) { 
  PropertiesService.getUserProperties().setProperty(INFO_URL_PROP, url);
  return true; 
}

/***** ================= USER CONTEXT (Via Admin Directory) ================= *****/
function getCurrentUserContext() {
  const email = Session.getActiveUser().getEmail();
  let fullName = "";
  let photoUrl = "";

  try {
    const user = AdminDirectory.Users.get(email, { viewType: 'domain_public' });
    if (user && user.name) {
      fullName = user.name.fullName || user.name.givenName;
    }
  } catch (e) {
    console.log("Admin Directory lookup failed: " + e.message);
  }

  if (!fullName) {
    const handle = email.split('@')[0];
    fullName = handle.split('.').map(part => part.charAt(0).toUpperCase() + part.slice(1)).join(' ');
  }

  return { email: email, name: fullName, photoUrl: photoUrl };
}

/***** ================= FORCE AUTHORIZATION ================= *****/
function authorizeScript() {
  DriveApp.getRootFolder(); 
  GmailApp.getInboxThreads(0, 1);
  
  // NEW: Forces Google to detect the MailApp permission requirement
  if (false) { MailApp.sendEmail("test@google.com", "test", "test"); }

  const tempDoc = DocumentApp.create("Insights Genie Permission Test");
  tempDoc.saveAndClose();
  DriveApp.getFileById(tempDoc.getId()).setTrashed(true);
  try {
    const email = Session.getActiveUser().getEmail();
    AdminDirectory.Users.get(email, { viewType: 'domain_public' });
  } catch (e) {
    console.warn('Admin Directory check failed: ' + e.message);
  }
}

/***** ================= CORE: Generate Custom Miscellaneous Insights ================= *****/
function generateMiscellaneousInsights(sheetUrl, targetaudience, customPrompt) {
  try {
    const sheetId = extractSheetId_(sheetUrl);
    if (!sheetId) throw new Error('Invalid Sheet URL');

    const ss = SpreadsheetApp.openById(sheetId);
    let miscSheet = ss.getSheetByName("Miscellaneous") || ss.getSheetByName("Misllenious");
    if (!miscSheet) throw new Error('Could not find a tab named "Miscellaneous" in the provided sheet.');

    const data = miscSheet.getDataRange().getDisplayValues();
    if (data.length < 2) throw new Error('No data found in the Miscellaneous tab.');

    const csvString = data.map(row => row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')).join('\n');

    const prompt = `ACT AS A SENIOR DATA ANALYST AND MARKET RESEARCHER.\nTOPIC: ${targetaudience}\nINSTRUCTIONS: ${customPrompt}\nDATA:\n${csvString}`;
    return { ok: true, reportText: callGeminiText_(prompt) };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function testMail() {
  MailApp.sendEmail(Session.getActiveUser().getEmail(), "Auth Test", "This is just to force the permission popup.");
}
