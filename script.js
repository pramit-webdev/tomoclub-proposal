document.addEventListener('DOMContentLoaded', () => {
    
    // ==========================================
    // DOM ELEMENTS
    // ==========================================
    const slides = document.querySelectorAll('.slide');
    const prevBtn = document.getElementById('prev-btn');
    const nextBtn = document.getElementById('next-btn');
    const progressFill = document.getElementById('progress-fill');
    const currentCounter = document.getElementById('current-slide');
    
    // View Toggles
    const viewSlidesBtn = document.getElementById('view-slides-btn');
    const viewReportBtn = document.getElementById('view-report-btn');
    const presentationContainer = document.getElementById('presentation-container');
    const submissionDocument = document.getElementById('submission-document');
    const printDocBtn = document.getElementById('print-doc-btn');
    
    // Presenter Portal Drawer
    const portalToggleBtn = document.getElementById('portal-toggle-btn');
    const portalCloseBtn = document.getElementById('portal-close-btn');
    const presenterPortal = document.getElementById('presenter-portal');
    const portalTabs = document.querySelectorAll('.portal-tab');
    const tabPanes = document.querySelectorAll('.tab-pane');
    
    const scriptContentText = document.getElementById('script-content-text');
    const notesContentText = document.getElementById('notes-content-text');
    
    // AI Speech Synthesis Player Controls
    const playVoiceBtn = document.getElementById('play-voice-btn');
    const stopVoiceBtn = document.getElementById('stop-voice-btn');
    const voiceSelect = document.getElementById('voice-select');
    const autoplayVoiceCheckbox = document.getElementById('autoplay-voice');
    
    // Interactive Prompt Builder (Slide 6)
    const gradeSelect = document.getElementById('grade-select');
    const subjectSelect = document.getElementById('subject-select');
    const formatSelect = document.getElementById('format-select');
    const promptOutputText = document.getElementById('prompt-output-text');
    const copyPromptBtn = document.getElementById('copy-prompt-btn');

    // Document Print Buttons
    const printLayoutBtn = document.getElementById('print-layout-btn');

    // ==========================================
    // APP STATE
    // ==========================================
    let currentIdx = 0;
    const totalSlides = slides.length;
    let isPortalOpen = false;

    // ==========================================
    // PRESENTATION DATA MAP (Loom Scripts & Notes)
    // ==========================================
    const slideData = {
        0: { // Slide 1: Title
            script: "<strong>Hello TomoClub Selection Committee!</strong> My name is Pramit Das and today I'm excited to walk you through my curriculum design and evaluation framework for a professional development workshop on <em>AI Use in Assessments and Personalisation</em>. In this short video, I will walk you through the slides, but more importantly, explain my pedagogical framework, active learning choices, and why this workshop represents a practical, high-impact curriculum. Let's begin!",
            notes: "Welcome participants as they enter the virtual or physical room. Display this title slide on screen. Play light, high-energy background music to set a professional tone. Tech check: Ensure screenshare is operating correctly and the chat is open."
        },
        1: { // Slide 2: Context
            script: "Why does this workshop matter? We start by addressing the core trilemma facing educators today. First, severe burnout: teachers spend upwards of <strong>10+ hours a week</strong> grading and lesson planning. Second, the personalization gap: with over 30 students per room, manual differentiation is virtually impossible. Finally, the academic integrity crisis: the panic over generative AI cheating has locked schools into a futile arms race using inaccurate plagiarism detectors. By framing AI as a supportive 'teacher co-pilot', we shift it from a threat into an active time-saver that amplifies teacher capacity and respects student agency.",
            notes: "Objective: Validate the teacher's current stress. Read through the three cards with an empathetic, supportive voice. Ask participants: <em>'Which of these three cards hits closest to home in your classroom today?'</em> and invite them to answer in the chat to encourage early interaction."
        },
        2: { // Slide 3: Overview
            script: "This workshop is designed as a 2-hour, highly practical professional development session targeted at K-12 educators, instructional designers, and school leaders. A 2-hour block is the cognitive sweet spot—long enough to execute two full hands-on sandboxes but short enough to prevent screen fatigue. It is built around an active 'sandwich' model: we keep presentation slides under 20 minutes, giving teachers the maximum time to build resources they can actually use in their classrooms the next day.",
            notes: "Walkthrough the workshop specifications table. Emphasize that the session style is 'Practical, Collaborative, and Reflective' and that teachers will walk away with real, designed resources, not just theory. Introduce TomoClub's active learning culture."
        },
        3: { // Slide 4: Objectives
            script: "Our learning objectives are intentionally modeled on Bloom's Revised Taxonomy. Under 'Understand', educators will grasp the pedagogical shift to process-oriented, AI-resistant assessment. Under 'Apply', they will design formative questions and multi-level reading passages using role-based prompt templates. Under 'Evaluate', they will learn the vital skill of 'Critical AI Auditing'—reviewing AI outputs for hallucinations, grade appropriateness, and hidden biases. This ensures a balanced approach that combines tech proficiency with critical awareness.",
            notes: "Breakdown the three objective categories. Point out that 'Critical AI Auditing' is the most important skill we teach—teachers should never copy-paste AI outputs blindly. Ask: <em>'How many of you have had an AI tool generate something that was completely inaccurate?'</em> to build consensus on the auditing objective."
        },
        4: { // Slide 5: Flow
            script: "Here is the 120-minute session timeline. We structure the workshop in six blocks. We open with a 15-minute Hook to check teacher perceptions. We then provide a 20-minute Mini-Input on prompt structures. This leads directly to our core interactive blocks: a 30-minute prompt battle on assessment design, and a 25-minute personalization sandbox. We spend 15 minutes debating academic integrity, and close with a 15-minute reflection and action planning block. This layout ensures educators are actively creating for the majority of the session, keeping engagement high.",
            notes: "Present the session roadmap visually. Clarify how the 'sandwich' structure works: Theory -> Practical Application -> Theory -> Practical Application -> Reflection. Suggest that participants grab a notebook or split their screen for the upcoming hands-on sandboxes."
        },
        5: { // Slide 6: Activity 1
            script: "Our first major activity is the 'Assessment Prompt Battle'. I structured this as a cooperative grade-level breakout activity. Instead of just lecturing on prompt engineering, we give teachers an interactive prompt builder. On the slide, they can choose their grade, topic, and format, build an optimized prompt, copy it, and run it. The task isn't just to generate a quiz; it is to critically audit it, write down three edits to correct vocabulary or align it with their context, and share it. This transforms prompting from a solo task into a collaborative game.",
            notes: "Demonstrate the Interactive Prompt Builder live. Select Grade 8, Science, and Rubric. Click copy. Instruct teachers: <em>'When you go into your grade breakouts, copy this baseline prompt, paste it into ChatGPT or Claude, and then focus 80% of your energy on *editing* and *polishing* the result to make it perfect for your actual kids. Upload your final polished draft to our shared Padlet!'</em>"
        },
        6: { // Slide 7: Activity 2
            script: "Activity 2 moves the focus to equity and personalized pathways. Teachers are challenged to differentiate a single core science concept for three highly realistic student profiles: Profile A, a struggling reader needing a lower Lexile level; Profile B, a multilingual learner needing dual-language glossaries; and Profile C, an advanced student needing high-rigor modeling extension. Educators use AI to scaffold these versions in minutes, demonstrating how AI makes scalable personalization a reality without removing the productive cognitive struggle students need to grow.",
            notes: "Introduce the three student profiles. Emphasize that personalization does not mean making the work 'easy'—it means removing the <em>reading barrier</em> while keeping the <em>scientific thinking rigor</em> high. Guide teachers to use tools like Diffit or Brisk for quick scaffolding."
        },
        7: { // Slide 8: Ethics
            script: "Now, we address academic integrity and responsible AI use. We explicitly tell teachers to abandon the 'arms race' of AI plagiarism detectors, which are inaccurate, easily bypassed, and heavily flagged against non-native speakers. Instead, we advocate for structural, process-oriented assessment: grading outlines and research drafts, incorporating short oral defenses, and using hyper-local prompts. We also establish a strict safety checklist: absolute student data privacy, human-in-the-loop oversight, and clear transparency guidelines for student AI use.",
            notes: "This is often a heated topic. Speak calmly and authoritatively. Share the news that open AI detectors produce false positives. Focus the discussion on <em>pedagogical workarounds</em> rather than surveillance. Ask: <em>'What is one way you can grade the *process* of learning in your subject, rather than just the final product?'</em>"
        },
        8: { // Slide 9: MEL Live
            script: "To evaluate if learning occurs, my Monitoring, Evaluation, and Learning framework uses synchronous checks during the session. We use a pre-session diagnostic Mentimeter poll to understand baseline comfort levels. During the session, we review the live collaborative breakout deliverables on Padlet, using a 3-star facilitator check (Standard? Role? Constraints?). We wrap up the live session with a Google Form Exit Ticket that measures self-reported confidence and objective mastery, aiming for an 85% active participation rate and an NPS score above 4.5.",
            notes: "Explain the live MEL indicators to school leaders. Show how the facilitator evaluates breakouts in real-time, allowing them to redirect groups that are getting stuck. Emphasize the importance of the final digital exit ticket."
        },
        9: { // Slide 10: MEL Async
            script: "To guarantee true learning transfer, teachers must submit a Classroom AI Adaptation Folder within 14 days of the session to receive their PD certification. They submit one co-created assessment, one differentiated accommodation set, and a critical Prompt Log reflecting on the specific revisions they made. Long-term, 30 to 90 days out, we evaluate success by tracking three key impact metrics: a workload reduction index aiming for over 3 hours saved per week, a personalization adoption rate of over 70%, and a positive shift toward process-oriented grading.",
            notes: "Detail the async portfolio submission. Assure teachers that they have already built 50% of the submission during the live breakouts today! Explain that the goal is not perfection, but demonstrating a reflective, critical co-pilot practice."
        },
        10: { // Slide 11: Format Adaptations
            script: "To fit school network schedules, this curriculum is fully modular and adaptable. The synchronous format relies on live breakouts, real-time chat checks, and peer critique. For asynchronous, self-paced delivery, we break the 120 minutes into 6 short interactive micro-videos with embedded check-in questions using Edpuzzle, and replace breakouts with structured discussion boards where teachers are required to peer-review two colleagues' prompts. This ensures the high-quality active learning of TomoClub is preserved across formats.",
            notes: "Walkthrough the comparison table. Explain that both sync and async formats are fully standards-aligned and lead to the exact same portfolio submission for certification, providing maximum scheduling flexibility for districts."
        },
        11: { // Slide 12: Closing
            script: "In conclusion, this curriculum embodies TomoClub's core vision: preparing students for a future we cannot predict by empowering teachers with tools they can trust. By positioning AI as a collaborative, ethical co-pilot, we reduce teacher burnout, make individual personalization accessible, and allow educators to focus on what they do best: building meaningful human relationships. Thank you so much for your time, and I look forward to working with TomoClub to design future-ready classrooms.",
            notes: "Congratulate participants on completing the workshop. Share the links to the async certification submission portal, the exit ticket, and the prompt library. Invite final questions and comments. Close with a warm, inspirational tone."
        }
    };

    // ==========================================
    // PRESENTATION ENGINE CONTROLS
    // ==========================================
    function updateSlides() {
        slides.forEach((slide, idx) => {
            slide.classList.remove('active', 'prev', 'next');
            if (idx === currentIdx) {
                slide.classList.add('active');
            } else if (idx < currentIdx) {
                slide.classList.add('prev');
            } else {
                slide.classList.add('next');
            }
        });

        // Update progress bar
        const progress = ((currentIdx + 1) / totalSlides) * 100;
        progressFill.style.width = `${progress}%`;
        
        // Update slide counter
        currentCounter.textContent = currentIdx + 1;

        // Button states
        prevBtn.style.opacity = currentIdx === 0 ? '0.3' : '1';
        prevBtn.style.pointerEvents = currentIdx === 0 ? 'none' : 'auto';
        
        nextBtn.style.opacity = currentIdx === totalSlides - 1 ? '0.3' : '1';
        nextBtn.style.pointerEvents = currentIdx === totalSlides - 1 ? 'none' : 'auto';

        // Dynamic Presenter Portal content swap
        if (slideData[currentIdx]) {
            scriptContentText.innerHTML = slideData[currentIdx].script;
            notesContentText.innerHTML = slideData[currentIdx].notes;
            
            // Stop old slide speech and autoplay if checked
            if (typeof stopSpeaking === 'function') {
                stopSpeaking();
                if (autoplayVoiceCheckbox && autoplayVoiceCheckbox.checked) {
                    setTimeout(() => {
                        speakScript();
                    }, 500);
                }
            }
        }
    }

    function nextSlide() {
        if (currentIdx < totalSlides - 1) {
            currentIdx++;
            updateSlides();
        }
    }

    function prevSlide() {
        if (currentIdx > 0) {
            currentIdx--;
            updateSlides();
        }
    }

    // Navigation Click Handlers
    nextBtn.addEventListener('click', nextSlide);
    prevBtn.addEventListener('click', prevSlide);

    // Keyboard Shortcuts Navigation
    document.addEventListener('keydown', (e) => {
        // Only run shortcuts if the presentation slides are active
        if (presentationContainer.classList.contains('active-view')) {
            if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'Enter') {
                e.preventDefault();
                nextSlide();
            } else if (e.key === 'ArrowLeft') {
                e.preventDefault();
                prevSlide();
            } else if (e.key.toLowerCase() === 'p') {
                e.preventDefault();
                togglePresenterPortal();
            }
        }
    });

    // Touch Support for Swipe
    let touchstartX = 0;
    let touchendX = 0;

    document.addEventListener('touchstart', e => {
        touchstartX = e.changedTouches[0].screenX;
    });

    document.addEventListener('touchend', e => {
        touchendX = e.changedTouches[0].screenX;
        handleGesture();
    });

    function handleGesture() {
        if (presentationContainer.classList.contains('active-view')) {
            if (touchendX < touchstartX - 50) nextSlide();
            if (touchendX > touchstartX + 50) prevSlide();
        }
    }

    // ==========================================
    // VIEW SWITCHING (Slides vs Printable Report)
    // ==========================================
    function showSlidesView() {
        viewReportBtn.classList.remove('active');
        viewSlidesBtn.classList.add('active');
        submissionDocument.style.display = 'none';
        submissionDocument.classList.remove('active-view');
        presentationContainer.style.display = 'flex';
        presentationContainer.classList.add('active-view');
        printDocBtn.style.display = 'none';
        updateSlides(); // Sync slide elements
    }

    function showReportView() {
        viewSlidesBtn.classList.remove('active');
        viewReportBtn.classList.add('active');
        presentationContainer.style.display = 'none';
        presentationContainer.classList.remove('active-view');
        submissionDocument.style.display = 'block';
        submissionDocument.classList.add('active-view');
        printDocBtn.style.display = 'flex';
        
        // Auto-close presenter portal drawer on report view
        if (isPortalOpen) {
            closePresenterPortal();
        }
    }

    viewSlidesBtn.addEventListener('click', showSlidesView);
    viewReportBtn.addEventListener('click', showReportView);

    // ==========================================
    // PRESENTER PORTAL DRAWER INTERACTIVE STATES
    // ==========================================
    function togglePresenterPortal() {
        if (isPortalOpen) {
            closePresenterPortal();
        } else {
            openPresenterPortal();
        }
    }

    function openPresenterPortal() {
        presenterPortal.classList.add('open');
        presentationContainer.classList.add('portal-open');
        portalToggleBtn.classList.add('highlight-btn');
        isPortalOpen = true;
    }

    function closePresenterPortal() {
        presenterPortal.classList.remove('open');
        presentationContainer.classList.remove('portal-open');
        portalToggleBtn.classList.remove('highlight-btn');
        isPortalOpen = false;
    }

    portalToggleBtn.addEventListener('click', togglePresenterPortal);
    portalCloseBtn.addEventListener('click', closePresenterPortal);

    // Presenter Portal Tab Switching
    portalTabs.forEach(tab => {
        tab.addEventListener('click', () => {
            portalTabs.forEach(t => t.classList.remove('active'));
            tab.classList.add('active');

            const targetTab = tab.getAttribute('data-tab');
            tabPanes.forEach(pane => {
                pane.classList.remove('active');
                if (pane.getAttribute('id') === targetTab) {
                    pane.classList.add('active');
                }
            });
        });
    });

    // ==========================================
    // INTERACTIVE PROMPT BUILDER LOGIC
    // ==========================================
    function updatePromptBuilder() {
        const grade = gradeSelect.value;
        const topic = subjectSelect.value;
        const format = formatSelect.value;

        let outputPrompt = `Act as an expert K-12 curriculum designer. Create a standards-aligned ${grade} assessment for ${topic}. The output must be ${format}. Ensure the reading level is age-appropriate and avoid generalities.`;

        promptOutputText.innerHTML = outputPrompt;
    }

    // Set change triggers
    gradeSelect.addEventListener('change', updatePromptBuilder);
    subjectSelect.addEventListener('change', updatePromptBuilder);
    formatSelect.addEventListener('change', updatePromptBuilder);

    // Clipboard copy action
    copyPromptBtn.addEventListener('click', () => {
        const textToCopy = promptOutputText.textContent;
        navigator.clipboard.writeText(textToCopy).then(() => {
            copyPromptBtn.textContent = '📋 Prompt Copied!';
            copyPromptBtn.classList.add('copied');
            
            setTimeout(() => {
                copyPromptBtn.textContent = '📋 Copy Prompt to Clipboard';
                copyPromptBtn.classList.remove('copied');
            }, 2000);
        }).catch(err => {
            console.error('Failed to copy prompt: ', err);
            alert('Failed to copy prompt to clipboard. You can select the text and copy manually.');
        });
    });

    // ==========================================
    // NATIVE AI VOICE SPEECH SYNTHESIS ENGINE
    // ==========================================
    let synth = window.speechSynthesis;
    let voices = [];
    let utterance = null;

    function populateVoices() {
        if (!synth) return;
        voices = synth.getVoices();
        
        // Save currently selected value if any
        const currentVal = voiceSelect.value;
        
        voiceSelect.innerHTML = '<option value="">Default AI Voice</option>';
        
        voices.forEach(voice => {
            // Filter standard english/natural sounding voices
            if (voice.lang.includes('en') || voice.lang.includes('US') || voice.lang.includes('GB')) {
                const option = document.createElement('option');
                option.textContent = `${voice.name} (${voice.lang})`;
                option.value = voice.name;
                voiceSelect.appendChild(option);
            }
        });
        
        if (currentVal) {
            voiceSelect.value = currentVal;
        }
    }

    // Initialize voice list
    populateVoices();
    if (synth && synth.onvoiceschanged !== undefined) {
        synth.onvoiceschanged = populateVoices;
    }

    function speakScript() {
        if (!synth) return;
        
        // Stop current speech
        synth.cancel();

        // Get script element content and strip HTML tags
        const rawText = scriptContentText.innerHTML;
        const cleanText = rawText.replace(/<[^>]*>/g, ''); // Strip all strong, em tags for perfect reading

        utterance = new SpeechSynthesisUtterance(cleanText);

        // Select chosen voice if selected
        const selectedVoiceName = voiceSelect.value;
        if (selectedVoiceName) {
            const selectedVoice = voices.find(v => v.name === selectedVoiceName);
            if (selectedVoice) utterance.voice = selectedVoice;
        }

        // Natural, comfortable pacing
        utterance.rate = 1.05; 
        utterance.pitch = 1.0;

        utterance.onstart = () => {
            playVoiceBtn.textContent = '⏸️ Pause';
            playVoiceBtn.classList.add('highlight-btn');
            stopVoiceBtn.style.display = 'inline-block';
        };

        utterance.onend = () => {
            resetPlayerButtons();
        };

        utterance.onerror = () => {
            resetPlayerButtons();
        };

        synth.speak(utterance);
    }

    function resetPlayerButtons() {
        playVoiceBtn.textContent = '▶️ Read Aloud';
        playVoiceBtn.classList.remove('highlight-btn');
        stopVoiceBtn.style.display = 'none';
    }

    function togglePlayPause() {
        if (!synth) return;
        if (synth.speaking) {
            if (synth.paused) {
                synth.resume();
                playVoiceBtn.textContent = '⏸️ Pause';
            } else {
                synth.pause();
                playVoiceBtn.textContent = '▶️ Resume';
            }
        } else {
            speakScript();
        }
    }

    function stopSpeaking() {
        if (synth) {
            synth.cancel();
            resetPlayerButtons();
        }
    }

    // Bind speech buttons
    playVoiceBtn.addEventListener('click', togglePlayPause);
    stopVoiceBtn.addEventListener('click', stopSpeaking);

    // ==========================================
    // PRINT & PDF EXPORT
    // ==========================================
    function triggerPrintDialog() {
        window.print();
    }

    printDocBtn.addEventListener('click', triggerPrintDialog);
    printLayoutBtn.addEventListener('click', triggerPrintDialog);

    // ==========================================
    // INITIALIZATION
    // ==========================================
    updateSlides();
    updatePromptBuilder();

});
