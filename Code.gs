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

    // 3. CONSTRUCT PROMPT WITH BOTH DATASETS
    const prompt = `ACT AS A SENIOR MARKET RESEARCH LEAD. Create a high-impact, structured insight report based on the provided search data.
    TOPIC: ${targetaudience}
    DATA DEFINITIONS:
    Topic: The keyword or term used to find this product.
    Searches (Ad Opps): The raw volume of interest. CRITICAL FILTER: IGNORE any query with fewer than 1,000 Searches. Do not include these in any analysis.
    YoY Growth: The percentage change in interest over the last year.
    RAW DATASET:
    (Top 100 of ${termBundle.totalRows} queries):
    ${JSON.stringify(termBundle.top, null, 2)}
    
    LOCATION DATA (Regional Breakdown):
    ${JSON.stringify(locationBundle, null, 2)}

    REPORT STRUCTURE:
    1. TOP 3 EMERGING MARKET TRENDS
    Identify overarching themes in the data (using only terms with >1,000 searches). For each trend, provide:
    Trend Name: A descriptive title for the shift.
    Metrics: Use [YoY Growth %] to justify the trend.
    Evidence: Cite specific queries from the data that prove this trend exists for ${targetaudience}.

    2. Share of Searches
    Using the search data, identify the presence of specific brands versus generic terms.
    Relative Share Calculation: Do NOT use raw volumes. Instead, calculate the "Relative Share" percentage within the competitive set found in the data (e.g., if Brand A, B, and C are found, Brand A's share = Brand A Volume / Total Volume of Brands A+B+C).
    Analysis: Highlight which brands are dominating the "Share of Mind" for ${targetaudience} and which are losing ground based on YoY growth.
    Output: Present the "Relative Share %" and "YoY Growth %" for the identified top brands.

    3. Brand X Attribute (Keep it qualitative instead of quantitative )
    
    4. STRATEGIC MARKETING RECOMMENDATIONS
    Provide actionable advice directly linked to the insights above (Internal Data only).
    Campaign Pivot: What should the brand highlight in their next campaign?
    Messaging Strategy: Suggest specific "hooks" or language based on consumer search intent.
    Winning Channel: Recommend where to deploy (e.g., Search, Social, or Educational content) based on whether the queries are high-intent or top-of-funnel.

    5. REGIONAL & GEOGRAPHIC ANALYSIS (INDIA)
    Data Source: STRICTLY use the LOCATION DATA for this section. Do not use Terms Data here.
    Regional Grouping: Aggregate the state-level data into the following zones:
    North: (e.g., Delhi, Haryana, Punjab, UP, Uttarakhand, HP, J&K, Ladakh)
    South: (e.g., Tamil Nadu, Karnataka, Kerala, Telangana, Andhra Pradesh)
    West: (e.g., Maharashtra, Gujarat, Rajasthan, Goa)
    East: (e.g., West Bengal, Bihar, Odisha, Jharkhand)
    Central: (e.g., Madhya Pradesh, Chhattisgarh)
    North East: (e.g., Assam, Meghalaya, Manipur, etc.)
    Analysis:
    Regional Share of Search: Calculate the total search volume for each region and display it as a percentage of the total national volume (Region Volume / Total Location Volume).
    Growth Hotspots: Identify which region has the highest aggregated YoY Growth.
    Output:
    Provide a breakdown of "Share of Search %" per Region.
    Highlight the top 1-2 individual states driving the highest growth.

    IMPORTANT:
    Do NOT start with meta sentences like "Of course..." or "Here is your report".
    Start directly from section 1.
    No conversational filler.
    STRICTLY PLAIN TEXT: Do NOT use Markdown formatting (no asterisks like ** or *).
    Do NOT use bolding syntax.
    Use uppercase for main headings (e.g., "TREND 1:") to distinguish them without bolding symbols.
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
    generationConfig: { temperature: 0.3, maxOutputTokens: 30000 }
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


Maintain the theme of the example below across all infographics, but use complete insights:
${reportText} which is generated to make infographics:

drive.web-frontend_20260202.13_p1
Folder highlights
Reports primarily analyze AI Courses and Cricket Fan Engagement trends, with some data on Education Degrees.

<!DOCTYPE html><html><head><script src="https://cdn.tailwindcss.com"></script></head><body style="background:#f8fafc; padding:24px;"><div style="max-width:1152px; margin:0 auto; background:white; border-radius:16px; padding:32px; box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1);"><div class="max-w-6xl mx-auto bg-white border border-slate-200 rounded-2xl shadow-sm p-6 md:p-8">
    <div class="font-sans text-slate-800 p-4 md:p-8">

    <header class="max-w-6xl mx-auto mb-10 text-center">
        <div class="inline-block bg-sky-600 text-white px-4 py-1 rounded-full text-sm font-semibold tracking-wide uppercase mb-3">Search Insights Report</div>
        <h1 class="text-4xl md:text-5xl font-extrabold text-slate-900 mb-4">Cricket Fan Engagement</h1>
        <p class="text-lg text-slate-600 max-w-3xl mx-auto">An analysis of emerging trends, search behavior, and market dynamics shaping the digital cricket landscape.</p>
    </header>

    <main class="max-w-6xl mx-auto space-y-12">

        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">1. Top 3 Emerging Market Trends</h2>
                    <p class="text-slate-400 mt-1">Fan interest is diversifying and intensifying at an explosive rate.</p>
                </div>
                <svg class="w-10 h-10 text-sky-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 7h8m0 0v8m0-8l-8 8-4-4-6 6"></path></svg>
            </div>
            
            <div class="p-6 md:p-8 space-y-8">
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-4 flex items-center">
                        <span class="w-2 h-8 bg-sky-500 rounded mr-3"></span>
                        The Explosion of Niche & Emerging Matchups
                    </h3>
                    <p class="text-slate-600 mb-4 text-sm max-w-4xl">Fan interest is rapidly broadening beyond traditional rivalries. Queries for women's cricket and non-marquee international tours are showing astronomical YoY growth, signaling a diversification of the core audience.</p>
                    <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
                        <div class="bg-sky-50 p-4 rounded-xl border border-sky-100">
                            <div class="text-xs text-sky-600 uppercase font-bold">indw vs sa w</div>
                            <div class="text-3xl font-bold text-sky-900 mt-1 flex items-center gap-2">+582,890% <span class="text-sm font-medium text-slate-500">YoY</span></div>
                            <div class="text-xs text-slate-500 mt-1">India Women vs South Africa Women</div>
                        </div>
                        <div class="bg-sky-50 p-4 rounded-xl border border-sky-100">
                            <div class="text-xs text-sky-600 uppercase font-bold">wi vs ind</div>
                            <div class="text-3xl font-bold text-sky-900 mt-1 flex items-center gap-2">+120,673% <span class="text-sm font-medium text-slate-500">YoY</span></div>
                            <div class="text-xs text-slate-500 mt-1">West Indies vs India</div>
                        </div>
                        <div class="bg-sky-50 p-4 rounded-xl border border-sky-100">
                            <div class="text-xs text-sky-600 uppercase font-bold">aus vs sa</div>
                            <div class="text-3xl font-bold text-sky-900 mt-1 flex items-center gap-2">+42,232% <span class="text-sm font-medium text-slate-500">YoY</span></div>
                            <div class="text-xs text-slate-500 mt-1">Australia vs South Africa</div>
                        </div>
                    </div>
                </div>
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-4 flex items-center">
                        <span class="w-2 h-8 bg-sky-500 rounded mr-3"></span>
                        Future-Focused Anticipation
                    </h3>
                    <p class="text-slate-600 mb-4 text-sm max-w-4xl">The fanbase is highly engaged and planning ahead. Significant search growth for future tournaments, well in advance, indicates a shift from in-the-moment consumption to long-term, anticipatory interest.</p>
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div class="bg-slate-50 p-4 rounded-xl border border-slate-200">
                            <div class="text-xs text-slate-600 uppercase font-bold">ipl 2025</div>
                            <div class="text-3xl font-bold text-slate-900 mt-1 flex items-center gap-2">+21,131% <span class="text-sm font-medium text-slate-500">YoY</span></div>
                        </div>
                        <div class="bg-slate-50 p-4 rounded-xl border border-slate-200">
                            <div class="text-xs text-slate-600 uppercase font-bold">ipl 2025 schedule</div>
                            <div class="text-3xl font-bold text-slate-900 mt-1 flex items-center gap-2">+3,872% <span class="text-sm font-medium text-slate-500">YoY</span></div>
                        </div>
                    </div>
                </div>
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-4 flex items-center">
                        <span class="w-2 h-8 bg-sky-500 rounded mr-3"></span>
                        Demand for Hyper-Granular, Real-Time Data
                    </h3>
                    <p class="text-slate-600 mb-4 text-sm max-w-4xl">Fans demand more than just a live score. The massive growth in specific, long-tail "scorecard" queries shows a sophisticated need for detailed, structured, and match-specific data.</p>
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div class="bg-slate-50 p-4 rounded-xl border border-slate-200">
                            <div class="text-xs text-slate-600 uppercase font-bold">mumbai indians vs gujarat titans match scorecard</div>
                            <div class="text-2xl font-bold text-slate-900 mt-1">+141,454% YoY</div>
                        </div>
                        <div class="bg-slate-50 p-4 rounded-xl border border-slate-200">
                            <div class="text-xs text-slate-600 uppercase font-bold">punjab kings vs rcb match scorecard</div>
                            <div class="text-2xl font-bold text-slate-900 mt-1">+69,776% YoY</div>
                        </div>
                    </div>
                </div>
            </div>
        </section>

        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">2. Consumer Insights & Behavioral Themes</h2>
                    <p class="text-slate-400 mt-1">Decoding the intent behind every search query.</p>
                </div>
                <svg class="w-10 h-10 text-emerald-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 20h5v-2a3 3 0 00-5.356-1.857M17 20H7m10 0v-2c0-.656-.126-1.283-.356-1.857M7 20H2v-2a3 3 0 015.356-1.857M7 20v-2c0-.656.126-1.283.356-1.857m0 0a5.002 5.002 0 019.288 0M15 7a3 3 0 11-6 0 3 3 0 016 0zm6 3a2 2 0 11-4 0 2 2 0 014 0zM7 10a2 2 0 11-4 0 2 2 0 014 0z"></path></svg>
            </div>
            
            <div class="p-6 md:p-8 grid grid-cols-1 lg:grid-cols-3 gap-8">
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-2 flex items-center">
                        <span class="w-2 h-8 bg-emerald-500 rounded mr-3"></span>
                        The "Second Screen" Score Seeker
                    </h3>
                    <p class="text-slate-600 mb-4 text-sm">While watching a match, fans use search for immediate, detailed data. Their intent is transactional and time-sensitive, seeking specific formats like scorecards for player stats and ball-by-ball details.</p>
                    <div class="space-y-2 text-sm">
                        <div class="bg-emerald-50 p-3 rounded-lg border border-emerald-100 font-mono">"live cricket score"</div>
                        <div class="bg-emerald-50 p-3 rounded-lg border border-emerald-100 font-mono">"ipl live score"</div>
                        <div class="bg-emerald-50 p-3 rounded-lg border border-emerald-100 font-mono">"[team] vs [team] match scorecard"</div>
                    </div>
                </div>
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-2 flex items-center">
                        <span class="w-2 h-8 bg-emerald-500 rounded mr-3"></span>
                        The Abbreviation-Driven Fan
                    </h3>
                    <p class="text-slate-600 mb-4 text-sm">Engaged fans use shorthand as their native digital language. This indicates deep familiarity and a preference for speed. Search volume for abbreviations consistently outpaces full-text names.</p>
                    <div class="space-y-2 text-sm">
                        <div class="bg-emerald-50 p-3 rounded-lg border border-emerald-100"><span class="font-semibold text-emerald-800">"ind vs eng"</span> has 2x+ volume of <span class="text-slate-500">"india vs england"</span></div>
                        <div class="bg-emerald-50 p-3 rounded-lg border border-emerald-100">Common usage: <span class="font-semibold text-emerald-800">rcb vs pbks, mi vs dc</span></div>
                    </div>
                </div>
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-2 flex items-center">
                        <span class="w-2 h-8 bg-emerald-500 rounded mr-3"></span>
                        The Brand-Loyal Navigator
                    </h3>
                    <p class="text-slate-600 mb-4 text-sm">A large user segment bypasses generic terms, using search as a direct navigation tool to their preferred brand destination. This shows powerful brand loyalty and habit formation.</p>
                    <div class="space-y-2 text-sm">
                        <div class="bg-emerald-50 p-3 rounded-lg border border-emerald-100"><span class="font-semibold text-emerald-800">"cricbuzz"</span> is the single highest-volume query in the dataset.</div>
                        <div class="bg-emerald-50 p-3 rounded-lg border border-emerald-100">"cricinfo" also shows significant brand-led search volume.</div>
                    </div>
                </div>
            </div>
        </section>

        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">3. Brand Share of Voice (SOV) Analysis</h2>
                    <p class="text-slate-400 mt-1">A near-monopoly in brand-led search intent.</p>
                </div>
                <svg class="w-10 h-10 text-violet-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 3.055A9.001 9.001 0 1020.945 13H11V3.055z"></path><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M20.488 9H15V3.512A9.025 9.025 0 0120.488 9z"></path></svg>
            </div>
            
            <div class="p-6 md:p-8 grid grid-cols-1 md:grid-cols-2 gap-8 items-center">
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Relative Share of Voice</h3>
                    <div class="space-y-4">
                        <div>
                            <div class="flex justify-between mb-1">
                                <span class="text-base font-medium text-violet-700">Cricbuzz</span>
                                <span class="text-sm font-medium text-violet-700">96.76%</span>
                            </div>
                            <div class="w-full bg-slate-200 rounded-full h-4"><div class="bg-violet-600 h-4 rounded-full" style="width: 96.76%"></div></div>
                        </div>
                        <div>
                            <div class="flex justify-between mb-1">
                                <span class="text-base font-medium text-slate-600">Cricinfo</span>
                                <span class="text-sm font-medium text-slate-600">3.24%</span>
                            </div>
                            <div class="w-full bg-slate-200 rounded-full h-4"><div class="bg-slate-400 h-4 rounded-full" style="width: 3.24%"></div></div>
                        </div>
                    </div>
                </div>
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Analysis</h3>
                    <p class="text-slate-600 text-sm">Cricbuzz exhibits near-total market dominance in brand-led search, making it the default destination for loyal users. While Cricinfo shows slightly higher YoY growth (+9.47% vs. +6.50%), its minimal share doesn't pose a significant competitive threat in the current search landscape.</p>
                </div>
            </div>
        </section>

        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">4. Strategic Marketing Recommendations</h2>
                    <p class="text-slate-400 mt-1">Capitalizing on high-growth segments and user intent.</p>
                </div>
                <svg class="w-10 h-10 text-amber-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9.663 17h4.673M12 3v1m6.364 1.636l-.707.707M21 12h-1M4 12H3m3.343-5.657l-.707-.707m2.828 9.9a5 5 0 117.072 0l-.548.547A3.374 3.374 0 0014 18.469V19a2 2 0 11-4 0v-.531c0-.895-.356-1.754-.988-2.386l-.548-.547z"></path></svg>
            </div>
            
            <div class="p-6 md:p-8 space-y-8">
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-2 flex items-center">
                        <span class="w-2 h-8 bg-amber-500 rounded mr-3"></span>
                        Campaign Pivot: Target Emerging Segments
                    </h3>
                    <p class="text-slate-600 text-sm">Shift focus to high-growth areas. Feature women's cricket prominently in campaigns, leveraging the 582,890% YoY search growth. Create dedicated content for non-traditional matchups (e.g., IND vs WI) to capture this expanding audience.</p>
                </div>
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-2 flex items-center">
                        <span class="w-2 h-8 bg-amber-500 rounded mr-3"></span>
                        Messaging Strategy: Speak the Fan's Language
                    </h3>
                    <p class="text-slate-600 mb-4 text-sm">Adopt consumer shorthand in all copy. Use abbreviations like "IND vs ENG" in ads, SEO, and social posts. Lead with speed and detail to match user intent.</p>
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-4 text-sm">
                        <div class="bg-amber-50 border-l-4 border-amber-400 p-4 rounded-r-lg">
                            <p class="font-bold text-amber-800">Example Hook 1:</p>
                            <p class="text-slate-700 mt-1 italic">"Fastest IND vs SA Scorecard, Ball by Ball"</p>
                        </div>
                        <div class="bg-amber-50 border-l-4 border-amber-400 p-4 rounded-r-lg">
                            <p class="font-bold text-amber-800">Example Hook 2:</p>
                            <p class="text-slate-700 mt-1 italic">"Your Real-Time RCB vs PBKS Match Centre."</p>
                        </div>
                    </div>
                </div>
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-2 flex items-center">
                        <span class="w-2 h-8 bg-amber-500 rounded mr-3"></span>
                        Winning Channel: A Two-Pronged Approach
                    </h3>
                    <div class="space-y-3 mt-4">
                        <div class="p-4 bg-slate-50 rounded-lg border border-slate-200">
                            <div class="font-bold text-slate-800">Search (SEO/SEM): Capture High-Intent Users</div>
                            <p class="text-sm text-slate-600 mt-1">Dominate the "Second Screen" moment. Optimize for long-tail "scorecard" keywords for every single match to win bottom-of-funnel users.</p>
                        </div>
                        <div class="p-4 bg-slate-50 rounded-lg border border-slate-200">
                            <div class="font-bold text-slate-800">Content/Social: Build Anticipation</div>
                            <p class="text-sm text-slate-600 mt-1">Capture the "Future-Focused" fan early. Create content hubs for "IPL 2025" with predictions, analysis, and schedule trackers to engage users at the top of the funnel.</p>
                        </div>
                    </div>
                </div>
            </div>
        </section>

        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">5. External Market Context</h2>
                    <p class="text-slate-400 mt-1">Macro trends fueling the search behavior.</p>
                </div>
                <svg class="w-10 h-10 text-rose-400" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M2.25 21h19.5m-18-18v18m10.5-18v18m6-13.5V21M6.75 6.75h.75m-.75 3h.75m-.75 3h.75m3-6h.75m-.75 3h.75m-.75 3h.75M9 21v-3.375c0-.621.504-1.125 1.125-1.125h3.75c.621 0 1.125.504 1.125 1.125V21" /></svg>
            </div>
            
            <div class="p-6 md:p-8 grid grid-cols-1 md:grid-cols-3 gap-8">
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-2">Digital Streaming Revolution</h3>
                    <p class="text-slate-600 text-sm">The shift from TV to OTT platforms (JioCinema, Hotstar) and fragmented broadcast rights increases the need for centralized, reliable score and information hubs that are always accessible.</p>
                </div>
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-2">Proliferation of Fantasy Sports</h3>
                    <p class="text-slate-600 text-sm">The rise of platforms like Dream11 transforms passive viewing into active, data-driven engagement. This creates a massive, continuous demand for in-depth player stats and real-time data, directly fueling scorecard searches.</p>
                </div>
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-2">Growth in Tier-2/3 Cities</h3>
                    <p class="text-slate-600 text-sm">Increased smartphone and data penetration is bringing millions of new fans online. This presents a major, underserved opportunity for platforms to offer scores, commentary, and analysis in regional languages.</p>
                </div>
            </div>
        </section>

        <section class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="bg-slate-900 text-white p-6 md:p-8 flex items-center justify-between">
                <div>
                    <h2 class="text-2xl font-bold">6. Regional & Geographic Analysis (India)</h2>
                    <p class="text-slate-400 mt-1">Identifying established markets and emerging growth hotspots.</p>
                </div>
                <svg class="w-10 h-10 text-teal-400" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M9 6.75V15m6-6v8.25m.5-10.5h-7a2.25 2.25 0 00-2.25 2.25v10.5a2.25 2.25 0 002.25 2.25h7.5a2.25 2.25 0 002.25-2.25v-10.5a2.25 2.25 0 00-2.25-2.25z" /></svg>
            </div>
            
            <div class="p-6 md:p-8 grid grid-cols-1 lg:grid-cols-2 gap-8">
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Regional Share of Search</h3>
                    <div class="grid grid-cols-2 gap-4 text-center">
                        <div class="bg-slate-50 p-4 rounded-xl border border-slate-200">
                            <div class="text-3xl font-bold text-teal-700">29.7%</div>
                            <div class="text-sm text-slate-600 font-semibold">South</div>
                        </div>
                        <div class="bg-slate-50 p-4 rounded-xl border border-slate-200">
                            <div class="text-3xl font-bold text-teal-700">24.8%</div>
                            <div class="text-sm text-slate-600 font-semibold">North</div>
                        </div>
                        <div class="bg-slate-50 p-4 rounded-xl border border-slate-200">
                            <div class="text-3xl font-bold text-teal-700">24.4%</div>
                            <div class="text-sm text-slate-600 font-semibold">West</div>
                        </div>
                        <div class="bg-slate-50 p-4 rounded-xl border border-slate-200">
                            <div class="text-3xl font-bold text-teal-700">13.9%</div>
                            <div class="text-sm text-slate-600 font-semibold">East</div>
                        </div>
                        <div class="bg-slate-50 p-4 rounded-xl border border-slate-200">
                            <div class="text-2xl font-bold text-slate-500">5.2%</div>
                            <div class="text-sm text-slate-500">Central</div>
                        </div>
                        <div class="bg-slate-50 p-4 rounded-xl border border-slate-200">
                            <div class="text-2xl font-bold text-slate-500">2.0%</div>
                            <div class="text-sm text-slate-500">North East</div>
                        </div>
                    </div>
                </div>
                <div>
                    <h3 class="text-lg font-semibold text-slate-700 mb-4">Growth Hotspots</h3>
                    <p class="text-slate-600 mb-4 text-sm">The East region is the primary national growth engine, bucking negative trends seen elsewhere. The most significant emerging state-level markets are in this region, representing the next wave of audience growth.</p>
                    <div class="space-y-4">
                        <div class="bg-teal-50 p-4 rounded-xl border border-teal-100">
                            <div class="text-xs text-teal-600 uppercase font-bold">Top Growth State</div>
                            <div class="text-2xl font-bold text-teal-900 mt-1">Bihar</div>
                            <div class="text-lg font-semibold text-green-600">+16.83% YoY</div>
                        </div>
                        <div class="bg-teal-50 p-4 rounded-xl border border-teal-100">
                            <div class="text-xs text-teal-600 uppercase font-bold">#2 Growth State</div>
                            <div class="text-2xl font-bold text-teal-900 mt-1">Assam</div>
                            <div class="text-lg font-semibold text-green-600">+9.63% YoY</div>
                        </div>
                    </div>
                </div>
            </div>
        </section>

    </main>

    <footer class="text-center text-sm text-slate-500 mt-12 pt-8 border-t border-slate-200">
        <p>Made by Samarth Sharma & Gemini</p>
    </footer>

    </div>
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

    const parsed = JSON.parse(resp.getContentText() || '{}');
    let htmlContent = parsed?.candidates?.[0]?.content?.parts?.[0]?.text || '';
    htmlContent = htmlContent.replace(/```html/gi, '').replace(/```/g, '').trim();

    const hostingResult = saveHtmlAndGetUrl_(targetaudience, htmlContent);
    return { ok: true, html: htmlContent, url: hostingResult.url };
  } catch (err) {
    return { ok: false, error: "Visual Hosting Error: " + err.message };
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
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e) {
    console.warn("Could not set public sharing permissions.");
  }

  const scriptUrl = ScriptApp.getService().getUrl();
  const finalUrl = scriptUrl.includes("/exec") ? `${scriptUrl}?view=${file.getId()}` : file.getUrl();

  return { url: finalUrl, fileId: file.getId() };
}

/***** ================= PUBLISH & EMAIL LOGIC ================= *****/
// Add this at the very top with your other constants
const LOGGING_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1ZfsNLTHFNitPQLxU1s0PHlR3y6z_DPtluldcZTEjyeU/edit';
function publishReportAndLog(audience, reportText, emailsCsv, person, client, rev) {
  try {
    const recipients = (emailsCsv || '').split(',').map(x => x.trim()).filter(Boolean);
    const userProps = PropertiesService.getUserProperties();
    const infoUrl = userProps.getProperty(INFO_URL_PROP) || '';

    // --- EXISTING LOGIC: Create Doc ---
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

    // --- EXISTING LOGIC: Set Permissions ---
    const file = DriveApp.getFileById(doc.getId());
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    recipients.forEach(r => { try { file.addViewer(r); } catch (err) {} });

    // ======================================================
    // 👇 THIS IS THE NEW PART THAT FIXES YOUR LOGGING 👇
    // ======================================================
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
    // ======================================================

    // --- EXISTING LOGIC: Send Email ---
    if (recipients.length) {
      let emailBody = `Here is your Insights Genie AI strategy report for ${audience}.\n\n`;
      emailBody += `📄 Strategy Doc: ${doc.getUrl()}\n`;
      if (infoUrl) emailBody += `📊 Visual Infographic: ${infoUrl}\n`;
      emailBody += `\nPrepared by: ${person} for ${client}`;

      GmailApp.sendEmail(recipients.join(','), `Insights Report: ${audience}`, emailBody);
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

// /***** ================= USER CONTEXT (Via People API) ================= *****/
// function getCurrentUserContext() {
//   const email = Session.getActiveUser().getEmail();
//   let fullName = "";

//   try {
//     // We ask People API for the 'names' field specifically
//     const person = People.People.get('people/me', { personFields: 'names' });
    
//     // Check if the API returned a valid name
//     if (person && person.names && person.names.length > 0) {
//       fullName = person.names[0].displayName;
//     }
//   } catch (e) {
//     console.log("People API Error: " + e.message);
//   }

//   // Fallback: If API fails or name is empty, use email handle
//   if (!fullName) {
//     const handle = email.split('@')[0];
//     fullName = handle.split('.')
//       .map(part => part.charAt(0).toUpperCase() + part.slice(1))
//       .join(' ');
//   }

//   return { email: email, name: fullName };
// }

// /***** ================= FORCE AUTHORIZATION ================= *****/
// function authorizeScript() {
//   console.log('1. Refreshing Drive Permissions...');
//   DriveApp.getRootFolder(); 
  
//   console.log('2. Refreshing Gmail Permissions...');
//   GmailApp.getInboxThreads(0, 1);

//   console.log('3. Refreshing Document Permissions...');
//   const tempDoc = DocumentApp.create("Insights Genie Permission Test");
//   tempDoc.saveAndClose();
//   DriveApp.getFileById(tempDoc.getId()).setTrashed(true);

//   // -------------------------------------------------------------
//   // CRITICAL CHANGE: NO TRY/CATCH HERE
//   // This forces the "Allow Access" popup to appear for Profile data
//   // -------------------------------------------------------------
//   console.log('4. Refreshing People API Permissions...');
//   const me = People.People.get('people/me', { personFields: 'names' });
  
//   // Log what it found to verify it works
//   if (me.names && me.names.length > 0) {
//     console.log("✅ Success! Found Name: " + me.names[0].displayName);
//   } else {
//     console.log("⚠️ API worked, but no Display Name found in your Google Profile.");
//   }

//   console.log('✅ Permissions fully updated.');
// }





/***** ================= USER CONTEXT (Via Admin Directory) ================= *****/
function getCurrentUserContext() {
  const email = Session.getActiveUser().getEmail();
  let fullName = "";
  let photoUrl = "";

  try {
    // 1. Try fetching via Admin Directory (Best for Workspace Domains)
    // viewType: 'domain_public' allows fetching basic info of users in the same domain
    const user = AdminDirectory.Users.get(email, { viewType: 'domain_public' });
    
    if (user) {
      // FIX: Prioritize fullName. If not available, use givenName (First Name)
      if (user.name) {
        fullName = user.name.fullName || user.name.givenName;
      }
      
    }
    
  } catch (e) {
    console.log("Admin Directory lookup failed (User might be outside domain or Service not enabled): " + e.message);
  }

  // 2. Fallback: If Admin SDK fails (e.g. personal @gmail.com), parse from email
  if (!fullName) {
    const handle = email.split('@')[0];
    fullName = handle.split('.')
      .map(part => part.charAt(0).toUpperCase() + part.slice(1)) // Title Case
      .join(' ');
  }

  // Return both name and photo
  return { email: email, name: fullName, photoUrl: photoUrl };
}
/***** ================= FORCE AUTHORIZATION ================= *****/
function authorizeScript() {
  console.log('1. Refreshing Drive Permissions...');
  DriveApp.getRootFolder(); 
  
  console.log('2. Refreshing Gmail Permissions...');
  GmailApp.getInboxThreads(0, 1);

  console.log('3. Refreshing Document Permissions...');
  const tempDoc = DocumentApp.create("Insights Genie Permission Test");
  tempDoc.saveAndClose();
  DriveApp.getFileById(tempDoc.getId()).setTrashed(true);

  // 4. Refreshing Admin Directory Permissions
  console.log('4. Refreshing Directory Permissions...');
  try {
    const email = Session.getActiveUser().getEmail();
    // specific call to trigger the scope prompt if not granted yet
    AdminDirectory.Users.get(email, { viewType: 'domain_public' });
    console.log('✅ Permissions fully updated.');
  } catch (e) {
    console.warn('⚠️ Admin Directory check failed (this is normal if using personal Gmail): ' + e.message);
  }
}




/***** ================= CORE: Generate Custom Miscellaneous Insights ================= *****/
function generateMiscellaneousInsights(sheetUrl, targetaudience, customPrompt) {
  try {
    const sheetId = extractSheetId_(sheetUrl);
    if (!sheetId) throw new Error('Invalid Sheet URL');

    const ss = SpreadsheetApp.openById(sheetId);
    
    // Look for the tab (checking a couple of spelling variations just in case)
    let miscSheet = ss.getSheetByName("Miscellaneous");
    if (!miscSheet) miscSheet = ss.getSheetByName("Misllenious");
    if (!miscSheet) throw new Error('Could not find a tab named "Miscellaneous" in the provided sheet.');

    // Fetch all data
    const data = miscSheet.getDataRange().getDisplayValues();
    if (data.length < 2) throw new Error('No data found in the Miscellaneous tab.');

    // Convert 2D array to CSV string. CSV is much more token-efficient than JSON for Gemini.
    const csvString = data.map(row => row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')).join('\n');

    const prompt = `ACT AS A SENIOR DATA ANALYST AND MARKET RESEARCHER.
    
    TOPIC/CONTEXT: ${targetaudience}
    
    USER'S CUSTOM INSTRUCTIONS:
    ${customPrompt}
    
    RAW DATASET (CSV format):
    ${csvString}
    
    OUTPUT RULES:
    1. Follow the user's custom instructions precisely.
    2. STRICTLY PLAIN TEXT: Do NOT use Markdown formatting (no asterisks like ** or *). 
    3. Use uppercase for main headings to distinguish them.
    4. Provide clear, actionable insights based ONLY on the data provided and the requested framing.`;

    const text = callGeminiText_(prompt);
    return { ok: true, reportText: text };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}
