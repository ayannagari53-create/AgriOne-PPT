/**
 * AgriOne Presentation Interactive Controller
 * Team: SILENT CODERS
 */

document.addEventListener("DOMContentLoaded", () => {
  const slides = Array.from(document.querySelectorAll(".slide"));
  const totalSlides = slides.length;
  let currentSlide = 1;

  // DOM Elements
  const slideCounter = document.getElementById("slideCounter");
  const categoryIndicator = document.getElementById("categoryIndicator");
  const progressBar = document.getElementById("progressBar");
  const slideSelect = document.getElementById("slideSelect");
  
  const btnPrev = document.getElementById("btnPrev");
  const btnNext = document.getElementById("btnNext");
  const btnFullscreen = document.getElementById("btnFullscreen");
  const btnOverview = document.getElementById("btnOverview");
  const btnNotes = document.getElementById("btnNotes");
  
  const drawerModal = document.getElementById("drawerModal");
  const btnCloseDrawer = document.getElementById("btnCloseDrawer");
  const drawerGrid = document.getElementById("drawerGrid");

  const notesModal = document.getElementById("notesModal");
  const btnCloseNotes = document.getElementById("btnCloseNotes");
  const notesTitle = document.getElementById("notesTitle");
  const notesBody = document.getElementById("notesBody");

  const timerBadge = document.getElementById("timerBadge");
  const timerText = document.getElementById("timerText");

  // Timer Variables
  let elapsedSeconds = 0;
  let timerInterval = setInterval(() => {
    elapsedSeconds++;
    const mins = String(Math.floor(elapsedSeconds / 60)).padStart(2, "0");
    const secs = String(elapsedSeconds % 60).padStart(2, "0");
    timerText.textContent = `${mins}:${secs}`;
  }, 1000);

  timerBadge.addEventListener("click", () => {
    elapsedSeconds = 0;
    timerText.textContent = "00:00";
  });

  // Comprehensive Speaker Notes for all 36 slides (Curated for 5-10 Min Pitch)
  const speakerNotes = {
    1: "Introduce Team Silent Coders with authority. Highlight Team Leader Iyaan Nagari and core members Bhuvan, Jafar, and Aiman. State the project mission: Building AgriOne as an intelligent digital agriculture platform.",
    2: "Establish AgriOne's core premise: We bridge farmers with AI and agronomic intelligence. Emphasize the core value chain: Farmer to Data to AI + Rule Engine to Insights to Better Decisions.",
    3: "Highlight modern farming pain points: Climate unpredictability, disease outbreaks, improper irrigation, soil uncertainty, and market exploitation. Note that these cause up to 40% yield loss and financial instability.",
    4: "Introduce the core philosophy: 'AI interprets, Rule Engine validates & calculates, AgriOne provides actionable decisions'. Clarify that we don't just chat—we compute scientifically safe actions.",
    5: "Walk through the complete 18-module ecosystem. AgriOne acts as a centralized command hub unifying everything from leaf scanning to livestock and drone intelligence.",
    6: "Showcase the Farmer Command Center. Point out how acreage, soil type, irrigation timers, and weather alerts are tailored specifically to the farmer's registered parcel.",
    7: "Explain geospatial precision: GPS pinpoints village coordinates, dynamically pulling localized weather and agro-climatic sub-zone agronomy.",
    8: "Demonstrate the AI Farmer Assistant. Emphasize that it answers the actual questions asked by the farmer with multi-domain agricultural context and safe chemical limits.",
    9: "Walk through AI Voice Interaction. In the field, farmers have dirt on their hands; voice-in and voice-out allows seamless hands-free advisory in regional accents.",
    10: "Address digital inclusivity. When a farmer selects their native language, the entire UI, audio synthesizer, and AI responses switch synchronously.",
    11: "Explain the AI Crop Health Scanner. Walk through the 6 steps: camera upload, neural vision detection, pathogen isolation, severity assessment, cause analysis, and dosage recommendation.",
    12: "Highlight Intelligent Water Management: Combining crop stage, soil type, moisture, and rainfall forecast to reduce water and electricity waste by up to 40%.",
    13: "Demonstrate weather intelligence translation: AgriOne turns raw mm and km/h into concrete directives (e.g., postpone fertilizer before 85% rain).",
    14: "Detail soil health diagnostics: Evaluating N-P-K, pH, and micronutrients to give customized organic and fertilizer blend recommendations per acre.",
    15: "Discuss crop recommendation: Matching seasonal cycles, soil water capacity, and historical Mandi market realization to answer 'What should I grow?'.",
    16: "Cover Mandi intelligence: Real-time price tracking across neighboring APMC markets prevents middleman exploitation and guides optimal selling windows.",
    17: "Showcase Digital Farm Management: Replacing lost paper notebooks with structured parcel records, crop lifecycles, and automated field reminders.",
    18: "Walk through the Daily Field Activity Checklist: Time-stamped daily operations keeping farm hands aligned and creating logs for organic certification.",
    19: "Explain Equipment & Machinery tracking: Engine runtime hours, fuel efficiency, and proactive maintenance alerts prevent disastrous harvest breakdowns.",
    20: "Explain the Farm-to-Market Supply Chain: Batch traceability, quality grading, and direct linkage with food processors eliminate 3-4 predatory middleman cuts.",
    21: "Introduce the Digital Marketplace: A dual-sided platform for selling farm produce at transparent rates and procuring certified authentic agricultural inputs.",
    22: "Cover Livestock & Dairy: Extending digital intelligence to dairy cattle, nutrition rationing, vaccination calendars, and milk productivity.",
    23: "Highlight Satellite & Drone Intelligence: High-altitude NDVI canopy monitoring combined with drone scouting reveals hidden crop stress without walking whole fields.",
    24: "Emphasize Sustainable Farming: Protecting depleting groundwater tables, preventing chemical over-spray, and rebuilding organic soil carbon.",
    25: "Walk through Agricultural Financial Intelligence: Farm ledger proof for low-interest bank crop loans and geo-tagged proof for rapid insurance claim settlements.",
    26: "Present Analytics & Insights: Visualizing yield gains (+18%), water savings (32%), and input cost reductions (-24%) across seasonal cycles.",
    27: "Explain Proactive Agricultural Alerts: Detect, Analyze, Alert, Act pipeline delivering critical early warnings via push, SMS, and regional voice calls.",
    28: "Review the complete 11-step user workflow: From login and GPS location setup to AI calculation and tangible farm execution.",
    29: "Walk through the enterprise technology architecture: Layered separation between frontend, Python FastAPI backend, AI layer, agronomic rule engine, and data stores.",
    30: "Emphasize why AgriOne is more than a chatbot: LLMs can hallucinate; our agricultural rule engine enforces strict mathematical limits and agronomy laws.",
    31: "Reiterate key differentiators: 18-module unified ecosystem, dual AI+Rule validation, voice-first accessibility, and hyper-local precision.",
    32: "Summarize impact: Tangible profitability and reduced uncertainty for individual farmers, paired with groundwater and soil conservation for the planet.",
    33: "Contrast Traditional Farming vs AgriOne: Moving from reactive intuition and fragmented hearsay to predictive, data-backed precision agronomy.",
    34: "Present the Future Scope: Clearly note upcoming phases including hardware IoT soil probes, autonomous drone sprayers, and predictive yield models.",
    35: "Deliver the closing vision: 'One Platform. Multiple Agricultural Intelligence Systems. One Smarter Farming Ecosystem.'",
    36: "Conclude with Thank You, re-introduce Silent Coders and Team Leader Iyaan Nagari, and invite questions from the evaluators."
  };

  // Populate Dropdown and Overview Drawer
  slides.forEach((slide, index) => {
    const slideNum = index + 1;
    const cat = slide.getAttribute("data-category") || "AgriOne";
    const titleEl = slide.querySelector(".slide-title, .hero-team-title, .concl-brand, .ty-title");
    const titleText = titleEl ? titleEl.textContent.trim() : `Slide ${slideNum}`;

    // Select Option
    const opt = document.createElement("option");
    opt.value = slideNum;
    opt.textContent = `${slideNum < 10 ? '0' + slideNum : slideNum}. [${cat}] ${titleText}`;
    slideSelect.appendChild(opt);

    // Overview Thumbnail Card
    const thumb = document.createElement("div");
    thumb.className = `thumb-card ${slideNum === 1 ? 'active' : ''}`;
    thumb.setAttribute("data-slide", slideNum);
    thumb.innerHTML = `
      <div class="thumb-num">SLIDE ${slideNum < 10 ? '0' + slideNum : slideNum} • ${cat}</div>
      <div class="thumb-title">${titleText}</div>
    `;
    thumb.addEventListener("click", () => {
      goToSlide(slideNum);
      closeOverview();
    });
    drawerGrid.appendChild(thumb);
  });

  // Navigation Logic
  function updateUI() {
    slides.forEach((slide, index) => {
      const num = index + 1;
      slide.classList.toggle("active", num === currentSlide);
    });

    slideCounter.textContent = `Slide ${currentSlide} / ${totalSlides}`;
    const activeSlide = slides[currentSlide - 1];
    const cat = activeSlide ? activeSlide.getAttribute("data-category") : "AGRIONE";
    categoryIndicator.textContent = cat.toUpperCase();

    // Progress Bar
    const pct = ((currentSlide) / totalSlides) * 100;
    progressBar.style.width = `${pct}%`;

    // Dropdown sync
    slideSelect.value = currentSlide;

    // Overview thumbnail sync
    const allThumbs = drawerGrid.querySelectorAll(".thumb-card");
    allThumbs.forEach(t => {
      t.classList.toggle("active", parseInt(t.getAttribute("data-slide")) === currentSlide);
    });

    // Update notes content if open
    updateNotesContent();
  }

  function goToSlide(num) {
    if (num < 1) num = 1;
    if (num > totalSlides) num = totalSlides;
    currentSlide = num;
    updateUI();
  }

  function nextSlide() {
    if (currentSlide < totalSlides) {
      goToSlide(currentSlide + 1);
    }
  }

  function prevSlide() {
    if (currentSlide > 1) {
      goToSlide(currentSlide - 1);
    }
  }

  function toggleFullscreen() {
    if (!document.fullscreenElement) {
      document.documentElement.requestFullscreen().catch(err => {
        console.warn("Fullscreen request error:", err);
      });
    } else {
      if (document.exitFullscreen) {
        document.exitFullscreen();
      }
    }
  }

  function openOverview() {
    drawerModal.classList.add("open");
  }

  function closeOverview() {
    drawerModal.classList.remove("open");
  }

  function updateNotesContent() {
    notesTitle.textContent = `Speaker Notes – Slide ${currentSlide} / ${totalSlides}`;
    const text = speakerNotes[currentSlide] || "Deliver the key concepts on this slide cleanly with focus on practicality and system impact.";
    notesBody.innerHTML = `
      <p><strong>🎯 Primary Presentation Objective:</strong></p>
      <p>${text}</p>
      <p style="margin-top: 16px; font-size: 13px; color: #94a3b8;">
        <em>💡 Pro Tip for Hackathon / Presentation: Speak clearly, focus on how AgriOne solves real farmer problems, and emphasize that AI is verified by the agricultural rule engine.</em>
      </p>
    `;
  }

  function toggleNotes() {
    notesModal.classList.toggle("open");
    if (notesModal.classList.contains("open")) {
      updateNotesContent();
    }
  }

  function closeNotes() {
    notesModal.classList.remove("open");
  }

  // Event Listeners
  btnNext.addEventListener("click", nextSlide);
  btnPrev.addEventListener("click", prevSlide);
  btnFullscreen.addEventListener("click", toggleFullscreen);
  btnOverview.addEventListener("click", openOverview);
  btnCloseDrawer.addEventListener("click", closeOverview);
  btnNotes.addEventListener("click", toggleNotes);
  btnCloseNotes.addEventListener("click", closeNotes);

  slideSelect.addEventListener("change", (e) => {
    goToSlide(parseInt(e.target.value));
  });

  // Global Keyboard Controls
  document.addEventListener("keydown", (e) => {
    // If overview or notes open, handle Esc
    if (e.key === "Escape") {
      closeOverview();
      closeNotes();
      return;
    }

    if (e.target.tagName === "INPUT" || e.target.tagName === "TEXTAREA" || e.target.tagName === "SELECT") {
      return;
    }

    switch (e.key) {
      case "ArrowRight":
      case "ArrowDown":
      case "PageDown":
      case " ":
        e.preventDefault();
        nextSlide();
        break;

      case "ArrowLeft":
      case "ArrowUp":
      case "PageUp":
        e.preventDefault();
        prevSlide();
        break;

      case "Home":
        e.preventDefault();
        goToSlide(1);
        break;

      case "End":
        e.preventDefault();
        goToSlide(totalSlides);
        break;

      case "f":
      case "F":
        e.preventDefault();
        toggleFullscreen();
        break;

      case "o":
      case "O":
        e.preventDefault();
        if (drawerModal.classList.contains("open")) {
          closeOverview();
        } else {
          openOverview();
        }
        break;

      case "n":
      case "N":
      case "p":
      case "P":
        e.preventDefault();
        toggleNotes();
        break;
    }
  });

  // Initial render
  updateUI();
});
