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

    // --- 3. PROCESS INTENT DATA ---
let intentBundle = [];
const intentSheet = ss.getSheetByName("Intent") || ss.getSheetByName("Intent_Data");
if (intentSheet) {
  const intentData = intentSheet.getDataRange().getValues();
  if (intentData.length > 1) {
    const iHeaders = intentData[0].map(h => h.toString().toLowerCase());
    const iName = iHeaders.findIndex(h => h.includes('intent') || h.includes('query') || h.includes('name'));
    const iOpp = iHeaders.findIndex(h => h.includes('opportunities') && !h.includes('pop'));
    const iPop = iHeaders.findIndex(h => h.includes('pop'));
    // NEW: Find the Share of Searches column
    const iShare = iHeaders.findIndex(h => h.includes('share of searches'));
    
    intentBundle = intentData.slice(1).map(r => ({
      intent: String(r[iName > -1 ? iName : 0]),
      volume: parseFloat(r[iOpp > -1 ? iOpp : 1]) || 0,
      growth: (parseFloat(r[iPop > -1 ? iPop : 2]) || 0) * 100,
      // Pass the pre-calculated share to Gemini (handle decimal vs percent)
      share: iShare > -1 ? (parseFloat(r[iShare]) > 1 ? parseFloat(r[iShare]) : parseFloat(r[iShare]) * 100) : null
    })).filter(r => r.intent);
  }
}

   // --- 4. PROCESS FEATURES DATA ---
let featureBundle = [];
const featureSheet = ss.getSheetByName("Features") || ss.getSheetByName("Features_Data");
if (featureSheet) {
  const featureData = featureSheet.getDataRange().getValues();
  if (featureData.length > 1) {
    const fHeaders = featureData[0].map(h => h.toString().toLowerCase());
    const fName = fHeaders.findIndex(h => h.includes('feature') || h.includes('query') || h.includes('name'));
    const fOpp = fHeaders.findIndex(h => h.includes('opportunities') && !h.includes('pop'));
    const fPop = fHeaders.findIndex(h => h.includes('pop'));
    const fShare = fHeaders.findIndex(h => h.includes('share of searches'));
    
    featureBundle = featureData.slice(1).map(r => ({
      feature: String(r[fName > -1 ? fName : 0]),
      volume: parseFloat(r[fOpp > -1 ? fOpp : 1]) || 0,
      growth: (parseFloat(r[fPop > -1 ? fPop : 2]) || 0) * 100,
      share: fShare > -1 ? (parseFloat(r[fShare]) > 1 ? parseFloat(r[fShare]) : parseFloat(r[fShare]) * 100) : null
    })).filter(r => r.feature);
  }
}

    // --- 5. PROCESS BRANDS DATA ---
let brandBundle = [];
const brandSheet = ss.getSheetByName("Brands") || ss.getSheetByName("Brands_Data");
if (brandSheet) {
  const brandData = brandSheet.getDataRange().getValues();
  if (brandData.length > 1) {
    const bHeaders = brandData[0].map(h => h.toString().toLowerCase());
    const bName = bHeaders.findIndex(h => h.includes('brand') || h.includes('query') || h.includes('name'));
    const bOpp = bHeaders.findIndex(h => h.includes('opportunities') && !h.includes('pop'));
    const bPop = bHeaders.findIndex(h => h.includes('pop'));
    const bShare = bHeaders.findIndex(h => h.includes('share of searches'));
    
    brandBundle = brandData.slice(1).map(r => ({
      brand: String(r[bName > -1 ? bName : 0]),
      volume: parseFloat(r[bOpp > -1 ? bOpp : 1]) || 0,
      growth: (parseFloat(r[bPop > -1 ? bPop : 2]) || 0) * 100,
      share: bShare > -1 ? (parseFloat(r[bShare]) > 1 ? parseFloat(r[bShare]) : parseFloat(r[bShare]) * 100) : null
    })).filter(r => r.brand);
  }
}


    const prompt = `ACT AS A SENIOR MARKET RESEARCH LEAD AND GOOGLE STRATEGIST. Create a high-impact, structured insight report based on the provided search data.
TOPIC: ${targetaudience}

DATA DEFINITIONS:
Topic: The keyword, feature, intent, or brand used to find this product.
Searches (Ad Opps): The raw volume of interest. CRITICAL FILTER: IGNORE any query with fewer than 1,000 Searches. Do not include these in any analysis.
YoY Growth: The percentage change in interest over the last year.
Share: The pre-calculated Share of Searches percentage.

RAW DATASET (General Terms):
(Top 100 of ${termBundle.totalRows} queries):
${JSON.stringify(termBundle.top, null, 2)}

INTENT DATA:
${JSON.stringify(intentBundle, null, 2)}

FEATURES DATA:
${JSON.stringify(featureBundle, null, 2)}

BRANDS DATA:
${JSON.stringify(brandBundle, null, 2)}

REPORT STRUCTURE:

EXECUTIVE ABSTRACT & AI DISCLAIMER
Provide a concise, 2-3 sentence high-level summary of the overall market dynamics for ${targetaudience}. Immediately follow this with a clear disclaimer stating: "Disclaimer: This report contains AI-generated strategic insights based on search data trends. Please independently verify all data before incorporating it into external or client-facing materials."

1. TOP 3 EMERGING MARKET TRENDS
Identify overarching themes in the data (using only terms with >1,000 searches). For each trend, provide:
Trend Name: A descriptive title for the shift.
Metrics: Use [YoY Growth %] to justify the trend.
Evidence: Cite specific queries from the data that prove this trend exists for ${targetaudience}.

2. Brand Share of Searches
Using the search data, identify the presence of specific brands versus generic terms.
Relative Share: Do NOT use raw volumes. Instead, strictly use the pre-calculated "share" values provided directly in the BRANDS DATA.
Analysis: Highlight which brands are dominating the Share of Mind for ${targetaudience} and which are losing ground based on YoY growth.
Format/Output: Present a clean breakdown showing the Share of Search (%) and the YoY Growth (%) for each brand. Add a brief qualitative summary of the competitive landscape (e.g., Who is the undisputed leader? Who are the rising challengers?).

3. The Intent Landscape
Analyze the "Intent" data to map where consumers are in the funnel.
Objective: strictly use the pre-calculated "share" values provided in the INTENT DATA to represent consumer intents.
Format: Present the Intent Share (%) and YoY Growth (%). Add commentary on shifting funnel dynamics (e.g., Is there a massive rise in "Comparison" or "Refurbished" queries indicating price sensitivity? Are "Review" searches outpacing "Best" searches?)

4. BRAND X ATTRIBUTE
STRICT DATA ISOLATION RULE: For this section, you are strictly forbidden from looking at the "Brands", "Intent", or "Features" data tabs. You MUST derive the Brand x Attribute resonance EXCLUSIVELY by analyzing the granular queries within the "Top terms" data.
Objective: Find granular queries in the top terms where a specific brand is searched alongside a specific feature or attribute (e.g., a high-volume query like "[Brand A] camera" vs "[Brand B] battery"). Provide strong qualitative commentary on which brands are successfully owning specific narratives or attributes in the minds of consumers.
Format: Identify 3-4 interesting intersections directly from the query strings. Explain the consumer perception based on the search volume and growth of those specific raw queries.

5. Feature Demand Breakdown
Analyze the "Features" data to show what hardware/software capabilities matter most.
Objective: Strictly use the pre-calculated "share" values provided in the FEATURES DATA. Include only the top 6 features by share percentage - aggregate the rest into an "Others" category.
Format: Present the Feature Share (%) and YoY Growth (%). Add qualitative insights on which features are now considered "table stakes" (high share, flat growth) versus "innovation drivers" (lower share, explosive growth).

6. Strategic Marketing Recommendations (Google Ads & Content)
Translate the above insights into actionable sales pitches for the client.
Objective: Provide 4-5 concrete marketing recommendations.
Format: Group recommendations by strategy.
Content Strategy: How should brands pivot their ad creatives, video narratives, and messaging pillars based on the Features and Intent data?
Google Ads Strategy: How should they structure their Google Ads campaigns? (e.g., Bidding aggressively on mid-funnel comparison terms via Search, using explosive feature trends as hooks in YouTube Shorts campaigns, or leveraging Performance Max to capture shifting brand allegiances).

IMPORTANT:
Do NOT start with meta sentences like "Of course..." or "Here is your report".
Start directly from the EXECUTIVE ABSTRACT & AI DISCLAIMER.
No conversational filler.
STRICTLY PLAIN TEXT: Do NOT use Markdown formatting (no asterisks like ** or *).
Do NOT use bolding syntax.
Use uppercase for main headings (e.g., "1. TOP 3 EMERGING MARKET TRENDS") to distinguish them without bolding symbols.
THRESHOLD RULE: STRICTLY ignore all data points with "Searches" < 1,000.
VOLUME RULE: Never output raw search volume numbers. Use "Share %" for groups/competitors and "YoY Growth %" elsewhere.
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
- Clearly display numbers which generating graphs.
- Label Wrapping: Logic must be included to wrap long labels to prevent them from overlapping.
- Like in the below example always keep the footer in the end.
- Always start with abstract in the starting.
- Also add the some explanation regarding the metric 
- Make sure that the infographics should contain all the insights while generating it.
- Always add the variety of chart wherever it make sense for each sections for eg table chart,pie hart,histogram ,bar chart etc.
- Horizontal Scroller: Applied to the "Emerging Market Trends" (Section 1) so users can swipe/scroll left to right through the trends.
- Horizontal Bar Chart: Applied to the "Brand Share" (Section 2) for a clean competitive landscape view.
- Donut Chart: Applied to "The Intent Landscape" (Section 3)
- Brand Resonance Matrix: Enhanced the "Brand x Attribute" (Section 4) into a more distinct grid format.
- Horizontal Bar Chart: Applied to the "Feature Demand Breakdown" (Section 5) to visualize how different hardware features stack up against each other.
- Action Cards: Upgraded the "Strategic Marketing Recommendations" (Section 6) to look like distinct, actionable cards.
- "Disclaimer: Insights are AI-generated. Please independently verify all data before incorporating it into external or client-facing materials." this line should always be at the footer of the infographics

Maintain the theme of the example below across all infographics, but use complete insights:
${reportText} which is generated to make infographics:


<!DOCTYPE html><html><head><script src="https://cdn.tailwindcss.com"></script></head><body style="background:#f8fafc; padding:24px;"><div style="max-width:1152px; margin:0 auto; background:white; border-radius:16px; padding:32px; box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1);"><div class="max-w-6xl mx-auto bg-white border border-slate-200 rounded-2xl shadow-sm p-6 md:p-8 font-sans text-slate-800">

    <!-- Header & Executive Abstract -->
    <header class="max-w-5xl mx-auto mb-12 text-center">
        <div class="inline-block bg-blue-600 text-white px-4 py-1 rounded-full text-sm font-semibold tracking-wide uppercase mb-4">Search Insights Report</div>
        <h1 class="text-4xl md:text-5xl font-extrabold text-slate-900 mb-4">The Smartphone Market</h1>
        <p class="text-lg text-slate-600">An analysis of consumer search trends, brand dynamics, and feature demand shaping the future of mobile technology.</p>
    </header>

    <main class="space-y-12">

        <!-- Executive Abstract -->
        <section class="bg-slate-50 border border-slate-200 rounded-xl p-6">
            <h2 class="text-xl font-bold text-slate-900 mb-3">Executive Abstract</h2>
            <p class="text-slate-700">The smartphone market is defined by a clear bifurcation: intense loyalty and forward-looking anticipation for dominant brands like Apple and Samsung, contrasted with a massive, growing consumer demand for value and cost-saving acquisition methods. The next competitive battleground is emerging around tangible, on-device AI capabilities and hardware-based privacy features, which are driving the most significant shifts in search behavior.</p>
        </section>

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
                    
                    <!-- Trend 1 Card: AI Feature Race -->
                    <div class="snap-center flex-shrink-0 w-11/12 md:w-2/3 lg:w-1/2 bg-slate-50 p-6 rounded-xl border border-slate-200">
                        <h3 class="text-lg font-semibold text-slate-800 mb-2">The AI Feature Race</h3>
                        <p class="text-slate-600 mb-4 text-sm">Queries for specific AI phone features are showing explosive growth, indicating a major shift from abstract interest to tangible product consideration.</p>
                        <div class="grid grid-cols-2 gap-4">
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"new galaxy ai phone"</div><div class="text-2xl font-bold text-green-600 mt-1">+17,732%</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"apple intelligence"</div><div class="text-2xl font-bold text-green-600 mt-1">+4,424%</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"samsung ai smartphone"</div><div class="text-2xl font-bold text-green-600 mt-1">+377%</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"circle to search"</div><div class="text-2xl font-bold text-green-600 mt-1">+821%</div></div>
                        </div>
                    </div>

                    <!-- Trend 2 Card: Perpetual Product Anticipation -->
                    <div class="snap-center flex-shrink-0 w-11/12 md:w-2/3 lg:w-1/2 bg-slate-50 p-6 rounded-xl border border-slate-200">
                        <h3 class="text-lg font-semibold text-slate-800 mb-2">Perpetual Product Anticipation</h3>
                        <p class="text-slate-600 mb-4 text-sm">Consumers are searching for unannounced, future-generation models at an unprecedented rate, signaling a highly engaged and forward-looking customer base.</p>
                        <div class="grid grid-cols-2 gap-4">
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"oppo reno 15"</div><div class="text-2xl font-bold text-green-600 mt-1">+109k%</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"vivo v60"</div><div class="text-2xl font-bold text-green-600 mt-1">+27k%</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"iphone 17"</div><div class="text-2xl font-bold text-green-600 mt-1">+1,398%</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"samsung s25"</div><div class="text-2xl font-bold text-green-600 mt-1">+987%</div></div>
                        </div>
                    </div>

                    <!-- Trend 3 Card: Privacy as a Feature -->
                    <div class="snap-center flex-shrink-0 w-11/12 md:w-2/3 lg:w-1/2 bg-slate-50 p-6 rounded-xl border border-slate-200">
                        <h3 class="text-lg font-semibold text-slate-800 mb-2">Privacy as a Searchable Feature</h3>
                        <p class="text-slate-600 mb-4 text-sm">Privacy is evolving from a general concern into a specific, hardware-based feature that consumers are actively searching for, creating a new competitive attribute.</p>
                        <div class="grid grid-cols-1 gap-4">
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"privacy display samsung"</div><div class="text-2xl font-bold text-green-600 mt-1">+314,600%</div></div>
                            <div class="bg-white p-3 rounded-lg border border-slate-200"><div class="text-xs text-slate-500 uppercase font-bold">"anti spy screen"</div><div class="text-2xl font-bold text-green-600 mt-1">+1,250%</div></div>
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
                    <p class="text-slate-400 mt-1">Apple and Samsung command the market, but key challengers are showing significant momentum.</p>
                </div>
                <svg class="w-10 h-10 text-indigo-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"></path></svg>
            </div>
            
            <div class="p-6 md:p-8 grid grid-cols-1 lg:grid-cols-5 gap-8">
                <div class="lg:col-span-3">
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Relative Share & YoY Growth</h3>
                    <div class="space-y-4 text-sm">
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Apple</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 27.49%;">27.5%</div></div><div class="w-20 text-right font-medium text-green-600">+13.34%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Samsung</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 18.35%;">18.4%</div></div><div class="w-20 text-right font-medium text-red-600">-4.08%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">vivo</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 13.54%;">13.5%</div></div><div class="w-20 text-right font-medium text-green-600">+18.60%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">OPPO</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 8.82%;">8.8%</div></div><div class="w-20 text-right font-medium text-green-600">+20.53%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Realme</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 7.04%;">7.0%</div></div><div class="w-20 text-right font-medium text-red-600">-2.63%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Redmi</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 4.90%;">4.9%</div></div><div class="w-20 text-right font-medium text-red-600">-25.52%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Google</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 2.35%;">2.4%</div></div><div class="w-20 text-right font-medium text-green-600">+27.61%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Huawei</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-indigo-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 0.48%;">0.5%</div></div><div class="w-20 text-right font-medium text-green-600">+79.32%</div></div>
                    </div>
                </div>
                <div class="lg:col-span-2 bg-slate-50 p-6 rounded-xl border border-slate-200">
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Analysis</h3>
                    <div class="space-y-3 text-sm text-slate-600">
                        <p><span class="font-bold text-slate-800">Duopoly at the Top:</span> Apple is the undisputed leader and still growing. Samsung holds a strong second but shows slight erosion.</p>
                        <p><span class="font-bold text-slate-800">Rising Challengers:</span> vivo and OPPO demonstrate healthy growth. Google and Huawei are the most notable risers, showing a resurgence in interest.</p>
                        <p><span class="font-bold text-slate-800">Losing Ground:</span> Legacy and specialized brands like Redmi and OnePlus are losing mindshare in the search landscape.</p>
                    </div>
                </div>
            </div>
        </section>

        <!-- Section 3: The Intent Landscape -->
        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">3. The Intent Landscape</h2>
                    <p class="text-slate-400 mt-1">Price is the dominant consideration, with explosive growth in alternative acquisition methods.</p>
                </div>
                <svg class="w-10 h-10 text-amber-400" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M21 21l-5.197-5.197m0 0A7.5 7.5 0 105.196 5.196a7.5 7.5 0 0010.607 10.607z" /></svg>
            </div>
            
            <div class="p-6 md:p-8 grid grid-cols-1 lg:grid-cols-2 gap-8 items-center">
                <div class="relative w-64 h-64 mx-auto">
                    <div class="absolute inset-0 rounded-full" style="background-image: conic-gradient(from 0deg, #f59e0b 0% 50.43%, #64748b 50.43% 64.07%, #fbbf24 64.07% 77.42%, #a3a3a3 77.42% 100%);"></div>
                    <div class="absolute inset-5 bg-white rounded-full flex flex-col items-center justify-center text-center">
                        <div class="text-4xl font-bold text-slate-800">50.4%</div>
                        <div class="text-sm text-slate-500 font-semibold uppercase tracking-wider">Price Intent</div>
                    </div>
                </div>
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Analysis</h3>
                    <p class="text-slate-600 text-sm mb-4">The consumer journey is overwhelmingly governed by cost. The most significant shift is the explosive growth in value-seeking behaviors. Traditional upper-funnel queries like "Best" and "Reviews" are declining, suggesting consumers are moving more quickly to commercial and cost-related queries.</p>
                    <div class="grid grid-cols-2 gap-x-6 gap-y-2 text-sm">
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-amber-500 mr-2"></span>Price: <span class="ml-auto font-medium text-green-600">+4.2%</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-slate-500 mr-2"></span>Best: <span class="ml-auto font-medium text-red-600">-11.9%</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-amber-400 mr-2"></span>Comparison: <span class="ml-auto font-medium text-green-600">+4.1%</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-slate-500 mr-2"></span>Reviews: <span class="ml-auto font-medium text-red-600">-5.6%</span></div>
                        <div class="flex items-center col-span-2 border-t pt-2 mt-2 border-slate-200 font-bold text-slate-800">Explosive Growth Drivers:</div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-green-500 mr-2"></span>Contract: <span class="ml-auto font-bold text-green-600">+187.2%</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-green-500 mr-2"></span>Trade-In: <span class="ml-auto font-bold text-green-600">+141.5%</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-green-500 mr-2"></span>Used: <span class="ml-auto font-bold text-green-600">+122.7%</span></div>
                        <div class="flex items-center"><span class="w-3 h-3 rounded-full bg-green-500 mr-2"></span>Deals: <span class="ml-auto font-bold text-green-600">+93.1%</span></div>
                    </div>
                </div>
            </div>
        </section>

        <!-- Section 4: Brand x Attribute Resonance Matrix -->
        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">4. Brand x Attribute Resonance Matrix</h2>
                    <p class="text-slate-400 mt-1">How consumers connect brands with specific concepts and features.</p>
                </div>
                <svg class="w-10 h-10 text-teal-400" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M9.568 3H5.25A2.25 2.25 0 003 5.25v4.318c0 .597.237 1.17.659 1.591l9.581 9.581c.699.699 1.78.872 2.607.33a18.095 18.095 0 005.223-5.223c.542-.827.369-1.908-.33-2.607L11.16 3.66A2.25 2.25 0 009.568 3z" /><path stroke-linecap="round" stroke-linejoin="round" d="M6 6h.008v.008H6V6z" /></svg>
            </div>
            
            <div class="p-6 md:p-8 grid grid-cols-1 md:grid-cols-2 gap-6">
                <div class="bg-slate-50 p-6 rounded-xl border border-slate-200">
                    <h3 class="text-lg font-semibold text-slate-800 mb-2">Samsung x AI Leadership</h3>
                    <p class="text-slate-600 text-sm">Samsung has successfully established a powerful narrative linking its brand directly to the concept of an "AI Phone." Consumers are searching for Samsung's specific implementation of it.</p>
                </div>
                <div class="bg-slate-50 p-6 rounded-xl border border-slate-200">
                    <h3 class="text-lg font-semibold text-slate-800 mb-2">Apple x Branded Intelligence</h3>
                    <p class="text-slate-600 text-sm">Apple is building a distinct association with its "Apple Intelligence" branding. Consumers perceive it as a core, device-dependent capability, not just a feature.</p>
                </div>
                <div class="bg-slate-50 p-6 rounded-xl border border-slate-200">
                    <h3 class="text-lg font-semibold text-slate-800 mb-2">Samsung x Hardware Privacy</h3>
                    <p class="text-slate-600 text-sm">Samsung is on the verge of owning the narrative around tangible privacy. The query "privacy display samsung" shows strong, specific user interest in a hardware-level solution.</p>
                </div>
                <div class="bg-slate-50 p-6 rounded-xl border border-slate-200">
                    <h3 class="text-lg font-semibold text-slate-800 mb-2">Samsung x Feature Usability</h3>
                    <p class="text-slate-600 text-sm">"Circle to Search" is resonating strongly enough to generate significant "how-to" search volume, proving the feature is a tool users are actively adopting and seeking to master.</p>
                </div>
            </div>
        </section>

        <!-- Section 5: Feature Demand Breakdown -->
        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">5. Feature Demand Breakdown</h2>
                    <p class="text-slate-400 mt-1">Screen technology is the new battleground as foundational specs become table stakes.</p>
                </div>
                <svg class="w-10 h-10 text-rose-400" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M10.5 6h9.75M10.5 6a1.5 1.5 0 11-3 0m3 0a1.5 1.5 0 10-3 0M3.75 6H7.5m3 12h9.75m-9.75 0a1.5 1.5 0 01-3 0m3 0a1.5 1.5 0 00-3 0m-3.75 0H7.5m9-6h3.75m-3.75 0a1.5 1.5 0 01-3 0m3 0a1.5 1.5 0 00-3 0m-9.75 0h9.75" /></svg>
            </div>
            
            <div class="p-6 md:p-8 grid grid-cols-1 lg:grid-cols-5 gap-8">
                <div class="lg:col-span-3">
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Relative Share & YoY Growth</h3>
                    <div class="space-y-4 text-sm">
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">5G</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-rose-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 40.27%;">40.3%</div></div><div class="w-20 text-right font-medium text-red-600">-18.16%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Screen</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-rose-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 11.92%;">11.9%</div></div><div class="w-20 text-right font-medium text-green-600">+101.87%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Camera</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-rose-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 8.70%;">8.7%</div></div><div class="w-20 text-right font-medium text-green-600">+8.76%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Color</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-rose-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 6.73%;">6.7%</div></div><div class="w-20 text-right font-medium text-green-600">+57.97%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Battery</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-rose-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 6.28%;">6.3%</div></div><div class="w-20 text-right font-medium text-green-600">+20.73%</div></div>
                        <div class="flex items-center gap-3"><div class="w-24 font-medium text-slate-800">Storage</div><div class="flex-1 bg-slate-100 rounded-full h-6"><div class="bg-rose-500 h-6 rounded-full flex items-center justify-end pr-2 text-white font-bold" style="width: 5.59%;">5.6%</div></div><div class="w-20 text-right font-medium text-green-600">+3.16%</div></div>
                    </div>
                </div>
                <div class="lg:col-span-2 bg-slate-50 p-6 rounded-xl border border-slate-200">
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Analysis</h3>
                    <div class="space-y-4 text-sm text-slate-600">
                        <div>
                            <div class="font-bold text-slate-800">Table Stakes vs. Innovation</div>
                            <p>"5G" is now an expectation, not a selling point. Core features like "Camera" and "Battery" remain evergreen concerns.</p>
                        </div>
                        <div>
                            <div class="font-bold text-slate-800">The Screen is the Star</div>
                            <p>The primary innovation driver is unequivocally the "Screen," with interest more than doubling. This suggests advancements in display tech are most exciting to consumers.</p>
                        </div>
                    </div>
                </div>
            </div>
        </section>

        <!-- Section 6: Strategic Recommendations -->
        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">6. Strategic Marketing Recommendations</h2>
                    <p class="text-slate-400 mt-1">Actionable strategies for Google Ads and Content to capitalize on market shifts.</p>
                </div>
                <svg class="w-10 h-10 text-cyan-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9.663 17h4.673M12 3v1m6.364 1.636l-.707.707M21 12h-1M4 12H3m3.343-5.657l-.707-.707m2.828 9.9a5 5 0 117.072 0l-.548.547A3.374 3.374 0 0014 18.469V19a2 2 0 11-4 0v-.531c0-.895-.356-1.754-.988-2.386l-.548-.547z"></path></svg>
            </div>
            
            <div class="p-6 md:p-8 grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                <!-- Card 1 -->
                <div class="p-5 bg-slate-50 rounded-lg border border-slate-200 flex flex-col">
                    <div class="font-bold text-slate-800 mb-2">1. Shift from "Specs" to "AI Experiences"</div>
                    <p class="text-sm text-slate-600 flex-grow">Create compelling content that demonstrates tangible, real-world benefits of on-device AI, such as "How Galaxy AI edits your photos" or "5 ways Apple Intelligence saves you time."</p>
                    <div class="mt-4 text-xs font-semibold text-cyan-700 uppercase">Content Strategy</div>
                </div>
                <!-- Card 2 -->
                <div class="p-5 bg-slate-50 rounded-lg border border-slate-200 flex flex-col">
                    <div class="font-bold text-slate-800 mb-2">2. Build a "Maximum Value" Content Hub</div>
                    <p class="text-sm text-slate-600 flex-grow">Address price sensitivity by creating a dedicated site section for "Trade-In Value Calculators," "Best Contract Deals," and "Certified Refurbished vs. Used Guides."</p>
                    <div class="mt-4 text-xs font-semibold text-cyan-700 uppercase">Content Strategy</div>
                </div>
                <!-- Card 3 -->
                <div class="p-5 bg-slate-50 rounded-lg border border-slate-200 flex flex-col">
                    <div class="font-bold text-slate-800 mb-2">3. Spotlight Screen Innovation</div>
                    <p class="text-sm text-slate-600 flex-grow">Capitalize on the 101.87% growth in "Screen" interest. Develop hero content showcasing advancements in display technology, especially emerging features like privacy displays.</p>
                    <div class="mt-4 text-xs font-semibold text-cyan-700 uppercase">Content Strategy</div>
                </div>
                <!-- Card 4 -->
                <div class="p-5 bg-slate-50 rounded-lg border border-slate-200 flex flex-col">
                    <div class="font-bold text-slate-800 mb-2">4. Capture AI-Consideration with PMax</div>
                    <p class="text-sm text-slate-600 flex-grow">Launch Performance Max campaigns themed around "AI Phone" and "Smart AI Features." Target audiences searching for competitor flagship models to intercept users comparing AI capabilities.</p>
                    <div class="mt-4 text-xs font-semibold text-cyan-700 uppercase">Google Ads</div>
                </div>
                <!-- Card 5 -->
                <div class="p-5 bg-slate-50 rounded-lg border border-slate-200 flex flex-col">
                    <div class="font-bold text-slate-800 mb-2">5. Dominate High-Growth Commercial Niches</div>
                    <p class="text-sm text-slate-600 flex-grow">Create highly-targeted Search campaigns for exploding queries. Bid aggressively on keywords like "trade in my phone for [Brand]" and "[Brand] phone contract deals."</p>
                    <div class="mt-4 text-xs font-semibold text-cyan-700 uppercase">Google Ads</div>
                </div>
            </div>
        </section>

    </main>

    <footer class="text-center text-sm text-slate-500 mt-12 pt-8 border-t border-slate-200">
        <p>Disclaimer: Insights are AI-generated. Please independently verify all data before incorporating it into external or client-facing materials.</p>
    </footer>

</div></div></body></html>

TITLE: ${targetaudience}
`;

    const payload = {
      contents: [{ role: 'user', parts: [{ text: prompt }]}],
      generationConfig: { temperature: 0.2, maxOutputTokens: 30000 }
    };

    const resp = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload), muteHttpExceptions: true, timeout: 60000
    });

    const responseText = resp.getContentText();
    let parsed = {};
    
    try {
      parsed = JSON.parse(responseText);
    } catch (e) {
      throw new Error("Invalid API response format. The Gemini service might be overloaded.");
    }

    if (parsed.error) {
      throw new Error(`Gemini API Error: ${parsed.error.message}`);
    }

    let htmlContent = parsed?.candidates?.[0]?.content?.parts?.[0]?.text || '';
    
    if (!htmlContent) {
      if (parsed?.candidates?.[0]?.finishReason === "SAFETY") {
        throw new Error("Content blocked by Google Safety Filters.");
      }
      throw new Error("Gemini returned an empty response. Try again.");
    }

    htmlContent = htmlContent.replace(/```html/gi, '').replace(/```/g, '').trim();

    const hostingResult = saveHtmlAndGetUrl_(targetaudience, htmlContent);
    return { ok: true, html: htmlContent, url: hostingResult.url };
    
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

/***** ================= CORE: Generate Custom Infographic HTML ================= *****/
function generateCustomInfographicHtmlFromInsights(targetaudience, reportText) {
  try {
    const key = PropertiesService.getScriptProperties().getProperty(GEMINI_KEY_PROP);
    const url = `${GEMINI_BASE}${GEMINI_MODEL}:generateContent?key=${key}`;

    let prompt;
    if (typeof getCustomInfographicPrompt === "function") {
      prompt = getCustomInfographicPrompt(targetaudience, reportText);
    } else {
      throw new Error("Custom infographic prompt function (getCustomInfographicPrompt) not found in info.gs.");
    }

    const payload = {
      contents: [{ role: 'user', parts: [{ text: prompt }]}],
      generationConfig: { temperature: 0.2, maxOutputTokens: 30000 }
    };

    const resp = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload), muteHttpExceptions: true, timeout: 60000
    });

    const responseText = resp.getContentText();
    let parsed = {};
    
    try {
      parsed = JSON.parse(responseText);
    } catch (e) {
      throw new Error("Invalid API response format. The Gemini service might be overloaded.");
    }

    if (parsed.error) {
      throw new Error(`Gemini API Error: ${parsed.error.message}`);
    }

    let htmlContent = parsed?.candidates?.[0]?.content?.parts?.[0]?.text || '';
    
    if (!htmlContent) {
      if (parsed?.candidates?.[0]?.finishReason === "SAFETY") {
        throw new Error("Content blocked by Google Safety Filters.");
      }
      throw new Error("Gemini returned an empty response. Try again.");
    }

    htmlContent = htmlContent.replace(/```html/gi, '').replace(/```/g, '').trim();

    const hostingResult = saveHtmlAndGetUrl_(targetaudience, htmlContent);
    return { ok: true, html: htmlContent, url: hostingResult.url };
    
  } catch (err) {
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
    file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e) {
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e2) {
      console.warn("Could not set sharing permissions. File remains private to the creator.");
    }
  }

  const scriptUrl = ScriptApp.getService().getUrl();
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

    const file = DriveApp.getFileById(doc.getId());
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    recipients.forEach(r => { try { file.addViewer(r); } catch (err) {} });

    try {
      if (LOGGING_SHEET_URL) {
        const logSheet = SpreadsheetApp.openByUrl(LOGGING_SHEET_URL).getSheets()[0];
        logSheet.appendRow([
          new Date(),
          person,
          client,
          audience,
          doc.getUrl(),
          infoUrl,
          recipients.join(', ')
        ]);
      }
    } catch (logErr) {
      console.log("Logging failed, but continuing execution: " + logErr);
    }

    if (recipients.length > 0) {
      const emailSubject = `Strategic Insights Report: ${audience} | Prepared for ${client}`;
      
      const emailBody = `Dear ${person},\n\n` +
        `Please find the strategic AI insights report for ${audience}, prepared specifically for ${client}.\n\n` +
        `This document contains a comprehensive analysis of the latest search trends, competitive dynamics, and actionable marketing recommendations based on consumer intent.\n\n` +
        `📄 Access the Strategy Document here: ${doc.getUrl()}\n\n` +
        `Best regards,\n` +
        `Team Genie`;

      MailApp.sendEmail({
        to: recipients.join(','),
        subject: emailSubject,
        body: emailBody,
        name: "Insights Genie",
        noReply: true
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

/***** ================= USER CONTEXT ================= *****/
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

function authorizeScript() {
  DriveApp.getRootFolder(); 
  GmailApp.getInboxThreads(0, 1);
  if (false) { MailApp.sendEmail("test@google.com", "test", "test"); }

  const tempDoc = DocumentApp.create("Insights Genie Permission Test");
  tempDoc.saveAndClose();
  DriveApp.getFileById(tempDoc.getId()).setTrashed(true);
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

/***** ================= CORE: Generate Custom Insights ================= *****/
function generateCustomInsightsFromSheet(sheetUrl, targetaudience) {
  try {
    const sheetId = extractSheetId_(sheetUrl);
    if (!sheetId) throw new Error('Invalid Custom Sheet URL');

    const ss = SpreadsheetApp.openById(sheetId);
    
    const sheet = ss.getSheets()[0]; 
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) throw new Error('No data found in the custom sheet.');

    const csvString = data.map(row => row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')).join('\n');

    let prompt;
    if (typeof getCustomPrompt === "function") {
        prompt = getCustomPrompt(targetaudience, csvString);
    } else {
        prompt = `ACT AS A SENIOR DATA ANALYST AND MARKET RESEARCHER.\nTOPIC: ${targetaudience}\nAnalyze the following data and provide a highly structured, strategic insight report.\nDATA:\n${csvString}`;
    }

    const text = callGeminiText_(prompt);
    return { ok: true, reportText: text };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}
