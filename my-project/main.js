const fs = require('fs');
const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    Header, Footer, AlignmentType, LevelFormat, ExternalHyperlink,
    HeadingLevel, BorderStyle, WidthType, ShadingType, PageNumber,
    PageBreak, TabStopType, TabStopPosition,
} = require('docx');

// ------------- helpers -------------
const BRAND = "2E5BBA";
const ACCENT = "C8102E";
const MUTE = "555555";
const BORDER = { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" };
const BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };

function P(text, opts = {}) {
    return new Paragraph({
        spacing: { after: 100 },
        ...opts,
        children: opts.children || [new TextRun({ text, ...(opts.run || {}) })],
    });
}
function H1(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 320, after: 160 },
        children: [new TextRun({ text })],
    });
}
function H2(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 240, after: 120 },
        children: [new TextRun({ text })],
    });
}
function H3(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_3,
        spacing: { before: 200, after: 80 },
        children: [new TextRun({ text })],
    });
}
function bullet(text, level = 0, runs = null) {
    return new Paragraph({
        numbering: { reference: "bullets", level },
        spacing: { after: 60 },
        children: runs || [new TextRun({ text })],
    });
}
function num(text, level = 0) {
    return new Paragraph({
        numbering: { reference: "numbers", level },
        spacing: { after: 60 },
        children: [new TextRun({ text })],
    });
}
function rich(parts) {
    // parts: array of { text, bold, italic, color }
    return new Paragraph({
        spacing: { after: 100 },
        children: parts.map(p => new TextRun(p)),
    });
}
function callout(label, body, color = BRAND) {
    return new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [9360],
        rows: [
            new TableRow({ children: [
                    new TableCell({
                        borders: {
                            top: { style: BorderStyle.SINGLE, size: 16, color },
                            bottom: BORDER, left: BORDER, right: BORDER,
                        },
                        width: { size: 9360, type: WidthType.DXA },
                        shading: { fill: "F4F7FB", type: ShadingType.CLEAR },
                        margins: { top: 160, bottom: 160, left: 200, right: 200 },
                        children: [
                            new Paragraph({
                                spacing: { after: 80 },
                                children: [new TextRun({ text: label, bold: true, color, size: 22 })],
                            }),
                            ...(Array.isArray(body) ? body : [new Paragraph({ children: [new TextRun({ text: body })] })]),
                        ],
                    }),
                ]}),
        ],
    });
}

function priorityTable(rows) {
    // rows: [{ issue, why, effort, impact }]
    const header = new TableRow({
        tableHeader: true,
        children: ["Issue", "Why it matters", "Effort", "Impact"].map((t, i) => new TableCell({
            borders: BORDERS,
            width: { size: [4200, 3360, 900, 900][i], type: WidthType.DXA },
            shading: { fill: BRAND, type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 120, right: 120 },
            children: [new Paragraph({ children: [new TextRun({ text: t, bold: true, color: "FFFFFF" })] })],
        })),
    });
    const body = rows.map(r => new TableRow({
        children: [
            new TableCell({
                borders: BORDERS,
                width: { size: 4200, type: WidthType.DXA },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: r.issue })] })],
            }),
            new TableCell({
                borders: BORDERS,
                width: { size: 3360, type: WidthType.DXA },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: r.why })] })],
            }),
            new TableCell({
                borders: BORDERS,
                width: { size: 900, type: WidthType.DXA },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: r.effort })] })],
            }),
            new TableCell({
                borders: BORDERS,
                width: { size: 900, type: WidthType.DXA },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: r.impact, bold: true })] })],
            }),
        ],
    }));
    return new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [4200, 3360, 900, 900],
        rows: [header, ...body],
    });
}

function fixTable(rows) {
    // rows: [{ where, current, fix }]
    const header = new TableRow({
        tableHeader: true,
        children: ["Where", "Currently says", "Should say"].map((t, i) => new TableCell({
            borders: BORDERS,
            width: { size: [2160, 3600, 3600][i], type: WidthType.DXA },
            shading: { fill: ACCENT, type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 120, right: 120 },
            children: [new Paragraph({ children: [new TextRun({ text: t, bold: true, color: "FFFFFF" })] })],
        })),
    });
    const body = rows.map(r => new TableRow({
        children: [
            new TableCell({
                borders: BORDERS,
                width: { size: 2160, type: WidthType.DXA },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: r.where, bold: true })] })],
            }),
            new TableCell({
                borders: BORDERS,
                width: { size: 3600, type: WidthType.DXA },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: r.current, italics: true, color: MUTE })] })],
            }),
            new TableCell({
                borders: BORDERS,
                width: { size: 3600, type: WidthType.DXA },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: r.fix, bold: true })] })],
            }),
        ],
    }));
    return new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2160, 3600, 3600],
        rows: [header, ...body],
    });
}

// ------------- content -------------
const children = [];

// Cover
children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 1600, after: 400 },
    children: [new TextRun({ text: "Elite Haulerz Moving", bold: true, size: 56, color: BRAND })],
}));
children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 200 },
    children: [new TextRun({ text: "Website Audit & Improvement Roadmap", size: 36 })],
}));
children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 200 },
    children: [new TextRun({ text: "elitehaulerz.com", italics: true, color: MUTE, size: 24 })],
}));
children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 1200 },
    children: [new TextRun({ text: "Focus: Design • Lead Generation • Local SEO", color: MUTE, size: 22 })],
}));
children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Prepared May 2026", color: MUTE, size: 20 })],
}));
children.push(new Paragraph({ children: [new PageBreak()] }));

// Executive summary
children.push(H1("Executive Summary"));
children.push(P("Elite Haulerz has a real, established business with a clear value proposition (locally owned, flat hourly rates, 200-mile radius, full-service residential, commercial, long-distance, and specialty moving). The site reflects that, but it's leaving money on the table in three specific ways:"));

children.push(rich([
    { text: "1. Credibility leaks. ", bold: true },
    { text: "Spelling errors, inconsistent claims (\"20+ years\" on the homepage vs. \"almost 10 years\" on the About page), broken navigation slugs, and wrong content on the Local Moving page (it describes commercial moves) all erode trust at the exact moment a prospect is deciding whether to call you." },
]));
children.push(rich([
    { text: "2. The lead form is upside-down. ", bold: true },
    { text: "The detailed quote form (move date, home size, residential vs. commercial) is buried on /contact/ — yet the most-trafficked page (the homepage) only asks for Name / Phone / Email plus a required \"How did you hear about us?\" field, which is friction with no benefit to the visitor. Move the qualifying form to the homepage and you'll capture more — and better — leads." },
]));
children.push(rich([
    { text: "3. Local SEO is generic, not local. ", bold: true },
    { text: "There are no city-specific landing pages (\"Movers in St. Joseph, MI\", \"Movers in South Bend, IN\", etc.), no embedded Google reviews, no local-business schema, and the title tags / H1s don't include the geographic terms people actually search for. Competitors who do these basics will outrank you in Google's Local Pack regardless of the quality of your service." },
]));

children.push(callout("The good news",
    [P("None of the issues are structural. The site has the right pages, the right services, and a credible brand. Most of what's below is rewriting copy, restructuring forms, and adding ~6–10 new local-SEO pages. A motivated developer or marketer can move through the top-priority list in 2–3 weeks of focused work.")]
));

children.push(new Paragraph({ children: [new PageBreak()] }));

// Top 10
children.push(H1("Top 10 Highest-Impact Fixes"));
children.push(P("Ordered by ROI: do these in this sequence."));

children.push(priorityTable([
    { issue: "Fix the wrong content on the Local Moving page (currently describes commercial moves)", why: "Visitors searching \"local movers\" land here and read about office relocations — they bounce.", effort: "S", impact: "High" },
    { issue: "Replace the homepage form with the qualifying form (move date, size, service type, from/to ZIP)", why: "Better leads, fewer dead-end calls, easier to give an accurate quote.", effort: "S", impact: "High" },
    { issue: "Fix all spelling and grammar errors site-wide (Commerical, Avilable, Southbend, St. Joesph, SW Michgian, Let Out, etc.)", why: "Single biggest credibility hit. People judge a moving company on whether it looks careful.", effort: "S", impact: "High" },
    { issue: "Reconcile the experience claim — pick one (\"Since 2015\" / \"10+ years\") and use it everywhere", why: "Inconsistent claims read as exaggeration and undermine all other claims.", effort: "S", impact: "Med" },
    { issue: "Fix the Long Distance nav link (currently /services/apartment-moving/, should be /services/long-distance/)", why: "Confusing URL hurts SEO ranking for \"long distance movers\" and looks broken.", effort: "S", impact: "Med" },
    { issue: "Add 6 city landing pages: St. Joseph MI, Niles MI, Benton Harbor MI, South Bend IN, Mishawaka IN, Michigan City IN", why: "Google's Local Pack rewards pages that match the searcher's city. This is the single biggest local-SEO lever.", effort: "M", impact: "High" },
    { issue: "Embed Google reviews on the homepage and contact page", why: "Social proof is the #1 conversion factor for service businesses. You already have a Yelp link — surface star ratings.", effort: "S", impact: "High" },
    { issue: "Tighten the homepage hero: one primary CTA, visible phone, trust badges, headline that names the service area", why: "First 5 seconds determine bounce vs. quote-request. Three equal-weight buttons split attention.", effort: "M", impact: "High" },
    { issue: "Add LocalBusiness + MovingCompany schema markup, plus correct title tags / meta descriptions for every page", why: "Tells Google exactly what you do, where, and your hours/phone/reviews. Drives Local Pack ranking.", effort: "M", impact: "High" },
    { issue: "Consolidate Residential Moving + Local Moving (they overlap heavily) into one strong \"Residential & Local Moving\" page", why: "Currently splits SEO authority across two thin pages competing for the same keywords.", effort: "M", impact: "Med" },
]));

children.push(new Paragraph({ children: [new PageBreak()] }));

// Section 1: Credibility / Copy errors
children.push(H1("1. Credibility & Copy Errors"));
children.push(P("These are the lowest-effort, highest-trust-impact fixes. Have someone proofread every page; below is what I caught from a single review."));

children.push(H2("Spelling and grammar"));
children.push(fixTable([
    { where: "Homepage section heading", current: "Residential and Commerical Moving", fix: "Residential and Commercial Moving" },
    { where: "Homepage section heading", current: "Elite Packing (Full-Service Avilable)", fix: "Elite Packing (Full-Service Available)" },
    { where: "Homepage city list & body copy", current: "St.Joesph, MI / Southbend, IN", fix: "St. Joseph, MI / South Bend, IN (with the space, correct spelling)" },
    { where: "Blog post title (homepage \"Latest Posts\")", current: "What Makes Us the Best Movers SW Michgian Has to Offer?", fix: "What Makes Us the Best Movers SW Michigan Has to Offer?" },
    { where: "Local Moving page H3", current: "Let Out Local Movers in Michigan Handle Your Commercial Moving", fix: "Let Our Local Movers in Michigan Handle Your Commercial Moving" },
    { where: "Homepage \"Proudly serving\" section", current: "well do it all at a reasonable cost", fix: "we'll do it all at a reasonable cost" },
    { where: "Homepage hero", current: "Elite   Haulerz, LLC. (extra spaces)", fix: "Elite Haulerz, LLC." },
    { where: "Homepage CTA section", current: "Book NoW (mixed caps)", fix: "Book Now" },
    { where: "Homepage section heading", current: "EMpowering your Saving Journey for your move", fix: "Empowering Your Savings on Moving Day" },
]));

children.push(H2("Inconsistent claims"));
children.push(P("These contradict each other on different pages and need to be reconciled to a single number used site-wide."));
children.push(bullet("Homepage: \"20+ years of Experience\""));
children.push(bullet("About page: \"For almost 10 years now, Elite Haulerz, LLC has been assisting…\""));
children.push(bullet("Local Moving page: \"Elite Haulerz Has Been Handling Corporate Moves Since 2015\""));
children.push(P("Recommendation: pick the one that's actually true (likely \"Since 2015\" for the LLC, or frame the founder's personal experience separately as \"15+ years of moving experience, founded Elite Haulerz in 2015\"). Use the same phrasing in every place."));

children.push(H2("Geographic copy that contradicts itself"));
children.push(rich([
    { text: "The homepage hero says you serve ", italics: true },
    { text: "\"Berrien County, MI, and St. Joseph County, IN\"", italics: true, bold: true },
    { text: ". Two problems: (a) the canonical Indiana county for South Bend is St. Joseph County, but most people don't know that — they know \"South Bend.\" (b) There is also a St. Joseph County in Michigan, which adds confusion. Spell it out the way customers think: ", italics: true },
    { text: "\"Serving Berrien County, MI, plus South Bend, Mishawaka, La Porte and the surrounding Northern Indiana area.\"", bold: true },
]));

children.push(new Paragraph({ children: [new PageBreak()] }));

// Section 2: Lead generation
children.push(H1("2. Lead Generation & Conversion"));
children.push(P("Your goal is quote requests. Everything on the homepage should serve that goal. Right now several elements work against it."));

children.push(H2("Problem: the homepage form is too thin"));
children.push(P("Current homepage form fields: Name, Phone, Email, How did you hear about us?* (required)."));
children.push(P("This produces low-quality leads (you still have to call back to ask move date, size, service) and adds friction with the \"how did you hear\" required field, which provides no value to the visitor and asks them to do work for you before they've gotten anything."));
children.push(H3("Recommended homepage form (in this order)"));
children.push(num("Move date (date picker, default to next 4 weeks)"));
children.push(num("Moving from (ZIP code)"));
children.push(num("Moving to (ZIP code)"));
children.push(num("Home size (Studio / 1BR / 2BR / 3BR / 4BR+ / Office / Specialty item)"));
children.push(num("Service needed (Loading help / Full move / Packing + move / Specialty only)"));
children.push(num("Name"));
children.push(num("Phone"));
children.push(num("Email"));
children.push(P("\"How did you hear about us?\" should be optional and ideally moved to the thank-you page after submit, when the visitor has already converted."));

children.push(H2("Problem: hero section has three competing CTAs"));
children.push(P("The hero currently shows three buttons of equal visual weight: \"Get a Free Estimate\", \"Our Moving Services!\", and \"Gallery\". This splits attention. Conversion best practice is one primary CTA above the fold."));
children.push(H3("Recommended hero structure"));
children.push(bullet("H1 with the geographic anchor: \"Stress-Free Movers in Southwest Michigan & Northern Indiana\""));
children.push(bullet("Subhead: 1 sentence, what makes you different (locally owned, flat hourly rates, fully insured)"));
children.push(bullet("Primary CTA button: \"Get My Free Quote\" (scrolls to form or opens modal)"));
children.push(bullet("Secondary, smaller link: \"Or call (269) 240-3518\" — click-to-call"));
children.push(bullet("Trust strip directly below: \"Licensed & Insured • FMCSA Registered • MDOT Registered • [★★★★★ X Google reviews]\""));

children.push(H2("Problem: no social proof above the fold"));
children.push(P("Reviews are mentioned (\"check our reviews on online sites\") but never shown. For a service that involves trusting strangers in your home, this is a huge missed conversion opportunity."));
children.push(bullet("Embed a Google reviews widget (e.g., Elfsight, Trustindex, or Reviews.io) on the homepage and the contact page"));
children.push(bullet("Pull 3 written testimonials with first name + city onto the homepage in a rotating carousel"));
children.push(bullet("Add review-request automation: text every customer 24 hours after the move asking for a Google review"));

children.push(H2("Problem: \"Quotes from $295\" is hidden"));
children.push(P("This price hook is great — it prequalifies leads and removes a major friction point. Right now it only appears on /contact/. Move it to the homepage hero or the form area, and add the asterisk explanation: \"*Minimum charge for a 2-mover, 2-hour local move\" so it doesn't feel like bait-and-switch."));

children.push(H2("Other conversion improvements"));
children.push(bullet("Add a \"How it works\" 3-step graphic: 1. Tell us about your move  2. Get a flat hourly quote  3. We do the heavy lifting"));
children.push(bullet("Add an FAQ section to the homepage (or link clearly to /faq) — addresses the \"will they damage my stuff?\" objection before it kills the lead"));
children.push(bullet("Replace placeholder-only form labels with proper labels (accessibility, conversion, autofill)"));
children.push(bullet("Add a sticky mobile CTA bar at the bottom of every page on phones: \"Call\" + \"Get Quote\""));
children.push(bullet("Add live chat or SMS — for moves, people often want to ask one quick question before filling out a form"));

children.push(new Paragraph({ children: [new PageBreak()] }));

// Section 3: Local SEO
children.push(H1("3. Local SEO"));
children.push(P("Local SEO for movers comes down to four things: (1) city-specific landing pages, (2) Google Business Profile, (3) review velocity, and (4) on-page schema. You're missing or under-investing in all four."));

children.push(H2("Build city-specific landing pages"));
children.push(P("This is the biggest single lever. When someone searches \"movers in St. Joseph MI\", Google ranks pages that contain that exact phrase, in the title, H1, URL, and body. Currently you have one homepage that lists 15 cities in a flat list — that ranks for almost none of them."));
children.push(H3("Recommended initial set (Phase 1, 6 pages)"));
children.push(bullet("/locations/movers-st-joseph-mi/"));
children.push(bullet("/locations/movers-benton-harbor-mi/"));
children.push(bullet("/locations/movers-niles-mi/"));
children.push(bullet("/locations/movers-south-bend-in/"));
children.push(bullet("/locations/movers-mishawaka-in/"));
children.push(bullet("/locations/movers-michigan-city-in/"));
children.push(H3("Phase 2 (months 2–3)"));
children.push(bullet("New Buffalo, Bridgman, Stevensville, Coloma, Kalamazoo, Portage, Paw Paw (MI)"));
children.push(bullet("Granger, La Porte, Elkhart (IN)"));

children.push(H3("Template for each city page"));
children.push(num("Title tag: \"Movers in [City], [State] | Elite Haulerz Moving\""));
children.push(num("Meta description: 150-char summary that includes the city + service + phone"));
children.push(num("H1: \"Local Movers in [City], [State]\""));
children.push(num("First 100 words: mention the city by name 2–3 times naturally, mention nearby landmarks or neighborhoods (\"from downtown to [neighborhood]\")"));
children.push(num("Distance/drive time from your Berrien Springs base"));
children.push(num("3 services blocks: Local moves, Long-distance from [City], Specialty moves"));
children.push(num("2–3 local testimonials (from customers in or near that city)"));
children.push(num("Embedded Google Map of the service radius"));
children.push(num("LocalBusiness schema with the city as the areaServed"));
children.push(num("CTA: same quote form used everywhere else"));
children.push(P("Critical: do not duplicate copy across these pages. Google will treat duplicate location pages as spam. Each page needs at least 400–600 words of unique content about that specific city."));

children.push(H2("Google Business Profile"));
children.push(P("Your GBP is the single biggest source of leads for any local moving company. If it's not optimized, no amount of website work will fix that."));
children.push(bullet("Verify your GBP and confirm hours, phone, address, and service area match the website exactly"));
children.push(bullet("Set categories to \"Mover\" (primary) + \"Moving and storage service\", \"Piano moving service\", \"Trucking company\""));
children.push(bullet("Upload 30–50 photos: trucks, team, before/after, gallery — and replace any stale ones quarterly"));
children.push(bullet("Post weekly updates (offers, jobs completed, tips) — Google rewards active profiles"));
children.push(bullet("Reply to every review within 48 hours, positive and negative"));
children.push(bullet("Set up Google's \"Get a quote\" messaging feature — go directly into your inbox"));

children.push(H2("Review velocity"));
children.push(P("Movers with 100+ Google reviews and a 4.7+ star average dominate the Local Pack. You should be asking every single customer for a review."));
children.push(bullet("Set up a review-request automation: text + email 24 hours after the move with a one-click Google review link"));
children.push(bullet("Goal: 5–10 new Google reviews per month"));
children.push(bullet("Display the live review count on the homepage (\"Rated 4.9 across 127 Google reviews\")"));

children.push(H2("On-page SEO basics"));
children.push(fixTable([
    { where: "Homepage <title>", current: "Elite Haulerz – Elite Moving Services", fix: "Movers in Southwest Michigan & Northern Indiana | Elite Haulerz" },
    { where: "Homepage meta description", current: "Missing", fix: "Locally owned movers serving Berrien County MI and Northern IN since 2015. Flat hourly rates, fully insured. Free quote: (269) 240-3518." },
    { where: "Local Moving <title>", current: "Local Moving – Elite Haulerz", fix: "Local Movers in Berrien County, MI | Elite Haulerz" },
    { where: "Long Distance URL", current: "/services/apartment-moving/", fix: "/services/long-distance-moving/ (301 redirect old slug)" },
    { where: "Image alt text", current: "Inconsistent (some descriptive, some generic filenames)", fix: "Every image: \"[what it shows] — Elite Haulerz movers in [city]\"" },
    { where: "H1 hierarchy", current: "Multiple H1s per page (\"Book NoW\", \"Long Distance\", \"Elite Packing\" all H1)", fix: "One H1 per page, the rest H2/H3" },
]));

children.push(H2("Schema markup to add"));
children.push(P("Schema is invisible code that tells Google exactly what your business is. A developer can add this in 2–3 hours and it can move you up the Local Pack in weeks."));
children.push(bullet("LocalBusiness / MovingCompany schema on every page (name, address, phone, hours, geo, areaServed, priceRange)"));
children.push(bullet("AggregateRating schema referencing your Google reviews"));
children.push(bullet("Service schema for each of the 6 services with a description and price range"));
children.push(bullet("FAQPage schema on the FAQ page (drives rich snippets in search results)"));
children.push(bullet("BreadcrumbList schema in your service pages"));

children.push(new Paragraph({ children: [new PageBreak()] }));

// Section 4: Design
children.push(H1("4. Design & Brand"));
children.push(P("The site reads as functional but dated. The good news is you don't need a full redesign to get most of the conversion benefit — you need a tighter visual hierarchy, fewer competing styles, and modern photography."));

children.push(H2("Visual hierarchy issues"));
children.push(bullet("Multiple H1s per page — pick one, demote the rest to H2"));
children.push(bullet("Inconsistent capitalization in headings (\"Book NoW\", \"EMpowering\", \"Commerical\") — pick one style (Title Case is most professional) and apply everywhere"));
children.push(bullet("Section headings are all the same size and weight, so the page reads as a wall of equally-important text — vary heading sizes (H2 should be visibly bigger than H3)"));
children.push(bullet("Body paragraphs are very long blocks of text — break into 2–3 sentence paragraphs, add subheadings, use pull-quotes for key claims"));

children.push(H2("Brand consistency"));
children.push(bullet("Define a brand color palette (primary, secondary, accent, neutral) and use only those colors site-wide"));
children.push(bullet("Pick 2 fonts max: one for headings, one for body. Stick to them."));
children.push(bullet("Buttons should all look the same: same color, same shape, same size for primary actions"));
children.push(bullet("Logo appears at different sizes/aspect ratios on different pages — standardize"));

children.push(H2("Photography"));
children.push(P("You have authentic photos of your team and trucks — that's a real asset. But they're presented as portrait phone-shots inside random rectangles. Improvements:"));
children.push(bullet("Get one or two professional, wide-aspect lifestyle shots of the team in branded shirts loading a truck (used in hero / about)"));
children.push(bullet("Standardize the gallery to a uniform aspect ratio (16:9 or 4:3, not mixed)"));
children.push(bullet("Add captions with city name (\"Move completed in St. Joseph, MI, March 2026\") — also helps SEO"));
children.push(bullet("Replace the dog mascot photo on the homepage with the team — the dog is charming on a Meet the Team page but hurts trust as a primary hero element"));

children.push(H2("Page speed & technical"));
children.push(bullet("The site runs Elementor 4.0.7. Current Elementor is 3.x branch (which is actually newer — 4.0.7 looks like an older release). Update WordPress + Elementor + theme for performance and security patches"));
children.push(bullet("Run the site through Google PageSpeed Insights and address Core Web Vitals (LCP, CLS, INP). Mobile speed is a Google ranking factor"));
children.push(bullet("Compress all images — most appear to be uncompressed PNG/JPG out of WordPress at 768x1024 or larger"));
children.push(bullet("Add proper Open Graph tags so the site previews well when shared on Facebook and in text messages"));

children.push(H2("Mobile experience"));
children.push(P("Most of your traffic is mobile. Specifically check:"));
children.push(bullet("Phone number is tap-to-call everywhere (it appears to be — confirm)"));
children.push(bullet("Form fields are large enough to tap with a thumb and use the right keyboard (numeric for phone/ZIP)"));
children.push(bullet("Sticky bottom bar with \"Call\" and \"Quote\" buttons"));
children.push(bullet("Hero text isn't cut off and CTA is visible without scrolling"));

children.push(new Paragraph({ children: [new PageBreak()] }));

// Section 5: Page-by-page
children.push(H1("5. Page-by-Page Notes"));

children.push(H2("Homepage (/)"));
children.push(bullet("Replace H1 \"Elite Haulerz, LLC.\" with a benefit-driven, geo-specific H1"));
children.push(bullet("Replace lightweight contact form with the qualifying form"));
children.push(bullet("Add review carousel and trust badges above the fold"));
children.push(bullet("Add \"How it works\" 3-step graphic"));
children.push(bullet("Replace the city wall-of-text with a clickable grid of city links (each going to the matching /locations/ page)"));
children.push(bullet("Move \"Quotes from $295\" to the hero"));

children.push(H2("Contact page (/contact/)"));
children.push(bullet("Drop the dropdown options that aren't actually services you sell (Ubox, Pods, ABF UPACK appear in the size dropdown — confusing if you don't service those, useful if you do)"));
children.push(bullet("Form labels are placeholder-only — add proper labels"));
children.push(bullet("Add map of service area as an embed instead of a static image (lets users zoom)"));
children.push(bullet("Confirm form submissions go somewhere reliable (email + CRM) and you reply within the promised 24 hours"));
children.push(bullet("Add a thank-you page with the next-steps explained — this is also where Google Ads conversion tracking fires"));

children.push(H2("Local Moving page (/services/local-moving/)"));
children.push(rich([
    { text: "CRITICAL: ", bold: true, color: ACCENT },
    { text: "the body copy describes commercial moves, not local moves. Sentences like \"Nothing is more challenging for a business than relocating to another building or office space\" and \"Elite Haulerz Has Been Handling Corporate Moves Since 2015\" do not belong on a Local Moving page. Either rewrite this page to be about local residential moves, or merge it into the Residential Moving page." },
]));
children.push(bullet("Also typo: \"Let Out Local Movers\" → \"Let Our Local Movers\""));

children.push(H2("Specialty Moving page (/services/specialty-moving/)"));
children.push(bullet("Page is too short — add specifics: piano moving (you have a whole blog post about it), antiques, gun safes, hot tubs, pool tables, art, equipment"));
children.push(bullet("Add pricing context (\"Most piano moves run $X–$Y depending on stairs and distance\")"));
children.push(bullet("Reference the piano blog post and link to it"));
children.push(bullet("Add a photo gallery of specialty items you've moved"));

children.push(H2("Long Distance / Apartment Moving page"));
children.push(bullet("URL is /services/apartment-moving/ but the nav calls it \"Long Distance\" — fix the URL slug to /services/long-distance-moving/ and 301 redirect the old one"));
children.push(bullet("Or split into two pages — \"Apartment Moving\" and \"Long-Distance Moving\" are different services and should rank for different keywords"));

children.push(H2("Residential Moving page"));
children.push(bullet("Overlaps heavily with Local Moving — consolidate into one strong page or differentiate clearly: Local = within 50 miles; Residential = home moves at any distance; Long-Distance = interstate"));

children.push(H2("Commercial Moving page"));
children.push(bullet("Should be the place where the current Local Moving copy actually lives — much of it is good for commercial"));
children.push(bullet("Add B2B specifics: weekend/after-hours moves, IT equipment handling, certificate of insurance available on request, hourly minimums"));

children.push(H2("FAQ page"));
children.push(bullet("Add at least 12–15 questions — \"Are you insured?\" \"What's the minimum?\" \"Do you charge for travel time?\" \"What if you damage something?\" \"Do you move pianos?\" \"How far in advance should I book?\" \"What payment methods?\" \"Do you provide boxes?\" — these are exactly what people are about to ask on the phone"));
children.push(bullet("Mark up with FAQPage schema so answers can appear directly in Google results"));

children.push(H2("Blog (/services/blog/)"));
children.push(bullet("URL slug is awkward — blog should live at /blog/, not /services/blog/"));
children.push(bullet("Three posts is a thin start — aim for 1 post per month minimum, focused on local search terms (\"Best time of year to move in Michigan\", \"How to move a piano in Berrien County\", \"Moving from South Bend to St. Joseph: what to know\")"));
children.push(bullet("Each post should link out to relevant service pages and city pages"));

children.push(new Paragraph({ children: [new PageBreak()] }));

// Section 6: Recommended copy
children.push(H1("6. Recommended Homepage Copy (Drop-In Replacement)"));
children.push(P("Below is a complete rewrite of the key homepage sections. You can use this as-is or adapt the voice. It's structured to maximize quote requests."));

children.push(H2("Hero section"));
children.push(callout("H1",
    [P("Stress-Free Movers in Southwest Michigan & Northern Indiana", { run: { bold: true, size: 32 } })]
));
children.push(callout("Subhead",
    [P("Locally owned, fully insured, flat hourly rates. Serving Berrien County, MI and South Bend, Mishawaka, La Porte and surrounding Northern IN since 2015.")]
));
children.push(callout("Primary CTA",
    [P("Get My Free Quote (button)", { run: { bold: true } })],
    ACCENT
));
children.push(callout("Secondary CTA",
    [P("Or call (269) 240-3518 — open 8am–8pm")]
));
children.push(callout("Trust strip (right under the CTA)",
    [P("★★★★★ Rated 4.9 on Google · Licensed & Insured · FMCSA & MDOT Registered · Quotes from $295")]
));

children.push(H2("Why Elite Haulerz section (replaces \"Empowering Your Saving Journey…\")"));
children.push(P("Three columns, each with an icon and 1 sentence:"));
children.push(bullet("Locally owned. No franchise fees, no middlemen — savings go to you, not a corporate office.", 0, [
    new TextRun({ text: "Locally owned. ", bold: true }),
    new TextRun({ text: "No franchise fees, no middlemen — savings go to you, not a corporate office." }),
]));
children.push(bullet("Flat hourly rates. The price you're quoted is the price you pay. No fuel surcharges, no \"long-carry\" fees.", 0, [
    new TextRun({ text: "Flat hourly rates. ", bold: true }),
    new TextRun({ text: "The price you're quoted is the price you pay. No fuel surcharges, no \"long-carry\" fees." }),
]));
children.push(bullet("Fully insured. FMCSA and MDOT registered. Your stuff is covered, end-to-end.", 0, [
    new TextRun({ text: "Fully insured. ", bold: true }),
    new TextRun({ text: "FMCSA and MDOT registered. Your stuff is covered, end-to-end." }),
]));

children.push(H2("Services section (replaces individual H1 blocks)"));
children.push(P("Four cards, each with a photo, title, 1-sentence description, and a \"Learn more\" link:"));
children.push(bullet("Residential & Local Moving — Studios to 5-bedroom homes, anywhere within 200 miles of Berrien Springs."));
children.push(bullet("Long-Distance Moving — Cross-state moves between Southwest Michigan and Northern Indiana, with timely delivery and transparent pricing."));
children.push(bullet("Commercial & Office Moving — Weekend and after-hours moves available. We minimize downtime and provide certificates of insurance on request."));
children.push(bullet("Specialty Moving — Pianos, gun safes, pool tables, antiques, art, and oversized items handled with the right equipment and the right team."));

children.push(H2("How it works (new section)"));
children.push(num("Tell us about your move. Fill out the form or call us — takes 2 minutes."));
children.push(num("Get your flat hourly quote. Same day, in writing."));
children.push(num("We do the heavy lifting. Clean-cut, uniformed, on time."));

children.push(H2("Service area (replaces wall-of-cities)"));
children.push(P("Single sentence + clickable city grid:"));
children.push(P("\"We serve all of Berrien County and the surrounding Southwest Michigan and Northern Indiana area. Click your city for a local quote:\"", { run: { italics: true } }));
children.push(P("Then a 3-column grid of city links going to each /locations/ page."));

children.push(H2("Closing CTA section"));
children.push(P("\"Ready for a stress-free move? Get your free quote in under 2 minutes — or call (269) 240-3518 to talk to a real person right now.\"", { run: { bold: true, size: 26 } }));

children.push(new Paragraph({ children: [new PageBreak()] }));

// Section 7: Roadmap
children.push(H1("7. 30 / 60 / 90 Day Roadmap"));

children.push(H2("Days 1–30: Stop the bleeding"));
children.push(P("Goal: fix everything that costs you trust today. Mostly copy and form changes — minimal developer time."));
children.push(bullet("Proofread every page; fix all spelling/grammar errors"));
children.push(bullet("Reconcile experience claim site-wide"));
children.push(bullet("Fix Local Moving page content"));
children.push(bullet("Replace homepage form with qualifying form"));
children.push(bullet("Tighten hero: one H1, one primary CTA, trust strip"));
children.push(bullet("Fix Long Distance URL slug + 301 redirect"));
children.push(bullet("Update title tags and meta descriptions on all 10 main pages"));
children.push(bullet("Verify and optimize Google Business Profile"));
children.push(bullet("Set up review-request automation"));

children.push(H2("Days 31–60: Build the local SEO engine"));
children.push(P("Goal: start ranking for city + service combinations."));
children.push(bullet("Build Phase 1 city pages (6 pages: St. Joseph MI, Benton Harbor MI, Niles MI, South Bend IN, Mishawaka IN, Michigan City IN)"));
children.push(bullet("Add LocalBusiness + AggregateRating + Service schema across the site"));
children.push(bullet("Embed Google reviews widget on homepage and contact page"));
children.push(bullet("Consolidate Residential + Local Moving into one strong page"));
children.push(bullet("Expand FAQ to 12–15 questions, add FAQPage schema"));
children.push(bullet("Publish 2 blog posts targeting local search terms"));

children.push(H2("Days 61–90: Polish and scale"));
children.push(P("Goal: tighten conversion and expand SEO footprint."));
children.push(bullet("Build Phase 2 city pages (10 more cities)"));
children.push(bullet("Add \"How it works\" graphic, before/after gallery, video testimonials if possible"));
children.push(bullet("Add sticky mobile CTA bar"));
children.push(bullet("Set up live chat or SMS contact"));
children.push(bullet("Get professional photography of team and trucks"));
children.push(bullet("Update Elementor / WordPress / theme; run PageSpeed audit and fix Core Web Vitals"));
children.push(bullet("Set up Google Ads conversion tracking + a small local Google Ads campaign for high-intent searches (\"movers near me\", \"local movers [city]\")"));

children.push(H2("Ongoing"));
children.push(bullet("1 new blog post per month, each targeting a local long-tail keyword"));
children.push(bullet("Aim for 5–10 new Google reviews per month"));
children.push(bullet("Weekly Google Business Profile post"));
children.push(bullet("Quarterly review of analytics: which pages drive quote requests? Double down on what works"));

children.push(new Paragraph({ children: [new PageBreak()] }));

// Final
children.push(H1("Next Steps"));
children.push(P("Pick whichever of the following you'd like to tackle first and I can produce it:"));
children.push(bullet("A complete homepage rewrite as ready-to-paste HTML"));
children.push(bullet("All 6 Phase 1 city landing pages, written and SEO-optimized"));
children.push(bullet("A new contact form spec (fields, validation, thank-you page) ready to hand to a developer"));
children.push(bullet("A Google Business Profile optimization checklist with exact text for every field"));
children.push(bullet("A schema markup file (JSON-LD) ready to drop into the site's <head>"));
children.push(bullet("A copy edit pass on every existing page (track changes Word doc you can hand to whoever maintains the site)"));

// ------------- build doc -------------
const doc = new Document({
    styles: {
        default: { document: { run: { font: "Calibri", size: 22 } } },
        paragraphStyles: [
            { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
                run: { size: 36, bold: true, color: BRAND, font: "Calibri" },
                paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 0 } },
            { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
                run: { size: 28, bold: true, color: "222222", font: "Calibri" },
                paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
            { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
                run: { size: 24, bold: true, color: BRAND, font: "Calibri" },
                paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 } },
        ],
    },
    numbering: {
        config: [
            { reference: "bullets",
                levels: [
                    { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
                        style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
                    { level: 1, format: LevelFormat.BULLET, text: "◦", alignment: AlignmentType.LEFT,
                        style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
                ]},
            { reference: "numbers",
                levels: [
                    { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
                        style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
                ]},
        ],
    },
    sections: [{
        properties: {
            page: {
                size: { width: 12240, height: 15840 },
                margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
            },
        },
        headers: {
            default: new Header({
                children: [new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [new TextRun({ text: "Elite Haulerz Website Audit", color: MUTE, size: 18 })],
                })],
            }),
        },
        footers: {
            default: new Footer({
                children: [new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                        new TextRun({ text: "Page ", color: MUTE, size: 18 }),
                        new TextRun({ children: [PageNumber.CURRENT], color: MUTE, size: 18 }),
                    ],
                })],
            }),
        },
        children,
    }],
});

Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync("/sessions/sharp-optimistic-pasteur/mnt/outputs/Elite_Haulerz_Website_Audit.docx", buffer);
    console.log("OK");
});