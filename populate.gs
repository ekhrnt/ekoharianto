/**
 * populate.gs
 * Run populateAllSheets() once to fill all tabs with Eko Harianto's data.
 * Jalankan dari Apps Script editor: pilih fungsi populateAllSheets → Run.
 */

function populateAllSheets() {
  populateProfile();
  populateExperience();
  populateProjects();
  populateSkills();
  populateTestimonials();
  populateAgentContext();
  SpreadsheetApp.getActive().toast('✅ All sheets populated!', 'Done', 5);
}

// ─── PROFILE ──────────────────────────────────────────────────────────────────
function populateProfile() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName('profile');
  if (!sheet) sheet = ss.insertSheet('profile');
  sheet.clearContents();

  const data = [
    ['name',      'Eko Harianto'],
    ['nickname',  'Kori'],
    ['title',     'AVP, Head of IT'],
    ['tagline',   'I build systems that scale — secure infrastructure, compliant processes, and AI-powered workflows that actually get used.'],
    ['bio1',      'I\'ve spent the last 10 years figuring out how to make IT work better — not just keep it running. What started in hospitality IT (building hotel infrastructure from zero) evolved into enterprise-scale work at Shopee, managing everything from Active Directory to endpoint security for 500+ users across multiple entities.'],
    ['bio2',      'These days I lead IT at MoneeInsure, a digital insurance company, where my focus splits between keeping things secure and compliant (ISO 27001, OJK, UU PDP) and building internal tools that actually solve problems — AI agents, Google Apps Script automations, compliance dashboards. I like the intersection where governance meets practical engineering.'],
    ['location',  'Jakarta, Indonesia'],
    ['available', 'true'],
    ['email',     'ekoharianto.id@gmail.com'],
    ['linkedin',  'https://www.linkedin.com/in/eko-harianto/'],
    ['whatsapp',  '6282300011736'],
    ['photo',     'https://drive.google.com/file/d/1E_XI_3gIciV2NWfngPSOoh2m1zcLekJZ/view?usp=sharing'],
    ['m1_val',    '10+'],
    ['m1_label',  'Years Experience'],
    ['m2_val',    '500+'],
    ['m2_label',  'Users Managed'],
    ['m3_val',    '11'],
    ['m3_label',  'Projects Built'],
  ];

  sheet.getRange(1, 1, data.length, 2).setValues(data);
}

// ─── EXPERIENCE ───────────────────────────────────────────────────────────────
function populateExperience() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName('experience');
  if (!sheet) sheet = ss.insertSheet('experience');
  sheet.clearContents();

  const headers = ['company', 'role', 'period', 'description', 'tags'];
  const rows = [
    [
      'MoneeInsure Indonesia',
      'AVP, Head of IT',
      'Feb 2023 – Present',
      'Running IT for a digital life insurance company — which means staying on top of security, compliance, and operations all at once. On the governance side: led ISO 27001:2013 → 2022 transition audit with zero major findings, managing OJK regulatory compliance and appointed as Data Protection Officer (DPO) for UU PDP. On the builder side: shipped internal AI tools (underwriting agent, custom GPT for claims), Google Apps Script automations for endorsements and mass notifications, and a compliance dashboard with automated email reminders. Also spearheaded the Dukcapil population data integration from end to end.',
      'IT Security,GRC,ISO 27001,AI/LLM,Google Apps Script,DPO,OJK Compliance,Automation'
    ],
    [
      'Shopee',
      'Sr. Service Delivery, Corporate IT',
      'Jan 2023 – Jul 2024',
      'Handled end-to-end IT for 500+ users across Shopee\'s enterprise environment — desktops, network, Active Directory, DNS, DHCP, IT asset lifecycle, and endpoint security. Also took on full IT setup for the new SealInsure Indonesia entity, from device provisioning to go-live.',
      'ITSM,IT Asset Management,Active Directory,Networking,Endpoint Security'
    ],
    [
      'Shopee',
      'Helpdesk Engineer, Corporate IT',
      'Jul 2020 – Jan 2023',
      'First-level IT support in one of Southeast Asia\'s fastest-moving tech companies. Handled tickets across departments, maintained Windows systems and Active Directory user accounts, and got involved in infrastructure upgrade projects. Good environment to move fast and learn fast.',
      'IT Service Delivery,Networking,Active Directory,Windows Administration'
    ],
    [
      'Goodrich Suites, ARTOTEL Portfolio',
      'Asst. IT Manager',
      'Jul 2019 – Jun 2020',
      'Managed full IT operations for a boutique hotel — network reconfiguration, NAS Synology + LDAP deployment, IT budgeting, and CapEx planning. Built a Telegram bot for real-time network monitoring that cut incident response time significantly. One of my first automation projects.',
      'Network Infrastructure,IT Management,IT Budgeting,Automation,Telegram Bot'
    ],
    [
      'eL Royale Hotel Yogyakarta',
      'IT Supervisor (Pre-Opening Team)',
      'Oct 2018 – Jun 2019',
      'Joined as part of the pre-opening IT team — basically built the whole IT infrastructure from scratch before the hotel opened. Set up Linux servers (Ubuntu + Proxmox), FreeNAS, the full network stack (routers, switches, DNS, DHCP, firewall), and Property Management Systems. Coordinated vendors and trained staff before handover.',
      'Linux,Proxmox,FreeNAS,Network Infrastructure,PMS,IT Management'
    ],
    [
      'Zest Hotel Yogyakarta (Swiss-belhotel International)',
      'IT & Graphic Design Supervisor',
      'Feb 2018 – Oct 2018',
      'Ran daily IT operations — network, PMS, equipment — while also handling graphic design for marketing: banners, posters, social media visuals. Two very different skill sets, but both needed doing.',
      'Network Infrastructure,PMS,Unifi,Graphic Design'
    ],
    [
      'Eastparc Hotel Yogyakarta',
      'Guest Service Agent',
      'Nov 2016 – Feb 2018',
      'Front-of-house guest services at a 5-star hotel. Handled reservations, check-in/out, and guest records via PMS. Good grounding in hospitality operations before moving fully into IT.',
      'PMS,Guest Service,Hospitality'
    ],
    [
      'Student Park Hotel Apartment',
      'Guest Service Officer',
      'Apr 2016 – Nov 2016',
      'Guest services and front desk at a hotel apartment — reservations, check-in/out, and day-to-day operations.',
      'PMS,Guest Service,Hospitality'
    ],
    [
      'THE 1O1 Yogyakarta Tugu Hotel',
      'Concierge (Daily Worker)',
      'Sep 2015 – Mar 2016',
      'Concierge and guest assistance at a boutique hotel in Yogyakarta.',
      'Guest Service,Hospitality'
    ],
    [
      'Hotel Tentrem',
      'Concierge (Daily Worker)',
      'Jul 2015 – Sep 2015',
      'Concierge services at one of Yogyakarta\'s leading luxury hotels.',
      'Guest Service,Hospitality'
    ],
    [
      'Neo Hotel Group',
      'Cook Helper Pastry (Daily Worker)',
      'May 2015 – Jul 2015',
      'Kitchen support in the pastry section — early days in the hospitality world.',
      'Hospitality,F&B'
    ],
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

// ─── PROJECTS ─────────────────────────────────────────────────────────────────
function populateProjects() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName('projects');
  if (!sheet) sheet = ss.insertSheet('projects');
  sheet.clearContents();

  const headers = ['number', 'year', 'title', 'description', 'tech'];
  const rows = [
    ['01', '2024', 'AI Underwriting Review Agent', 'Built an AI agent on Quantum (internal AI platform) that automatically reviews incoming quotations from BD and external partners — assessing risk, coverage, and alignment with underwriting criteria. Cut manual review time significantly and brought more consistency to UW decisions.', 'AI/LLM,Quantum Platform,Underwriting Automation,Risk Assessment'],
    ['02', '2024', 'Endorsement Ticketing System', 'A lightweight ticketing system connecting Point of Sale (POS) and Customer Service for handling policyholder endorsement requests — built entirely in Google Apps Script with AI-assisted routing. Zero additional infrastructure cost; went from scattered requests to structured, trackable workflows.', 'Google Apps Script,AI Integration,Workflow Automation,Insurance Ops'],
    ['03', '2024', 'Mass Personalized Notifications', 'Built a system to send endorsement documents to thousands of policyholders — each message uniquely personalized with their name, policy ID, and product-specific content. Auto-generated and dispatched via notification pipeline, fully built in Google Apps Script.', 'Google Apps Script,Mass Notification,Personalization,Insurance Ops'],
    ['04', '2023', 'ISO Compliance Dashboard', 'A compliance tracker with document lifecycle calendar — monitors review and renewal deadlines for every controlled document, then automatically emails each auditee with reminders timed to their specific schedule. Audit readiness, automated.', 'Google Apps Script,ISO 27001,Email Automation,Google Workspace'],
    ['05', '2024', 'Custom GPT for Claims Assessment', 'Fine-tuned and deployed a custom GPT aligned to the internal claims assessment framework. Claims officers now use it to evaluate incoming claims faster and more consistently — AI-assisted decisions grounded in company policy and OJK insurance regulations.', 'Custom GPT,AI Integration,Claims Automation,Insurance Operations'],
    ['06', '2023', 'Inbound Number Migration', 'End-to-end ownership of migrating the company\'s customer service inbound number: requirements gathering, provider coordination, system reconfiguration, UAT, and go-live. Zero service disruption. All telephony and CRM dependencies covered.', 'Change Management,Telecoms,System Integration,IT Operations'],
    ['07', '2024', 'ISO 27001 Transition Audit', 'Led the ISO 27001:2013 → 2022 surveillance and transition audit at MoneeInsure. Outcome: "recommended" with zero major findings. Covered gap assessment, control updates, documentation overhaul, and coordinating with the external auditor.', 'ISO 27001,GRC,Audit Management,Policy Development'],
    ['08', '2023', 'Dukcapil Data Integration', 'End-to-end integration with Indonesia\'s Dukcapil (civil registry) system — coordinating legal, vendor, and infrastructure teams. Ensured compliant and secure population data exchange aligned with UU PDP requirements.', 'API Integration,Data Privacy,UU PDP,Vendor Management'],
    ['09', '2019', 'Hotel IT Infrastructure Build', 'Built the entire IT stack for eL Royale Hotel Yogyakarta before opening day — Linux servers (Ubuntu + Proxmox), FreeNAS NAS, full network (routers, switches, DNS, DHCP, firewall), PMS deployment, vendor coordination, and staff training. Zero to ready.', 'Linux,Proxmox,FreeNAS,Network Design,PMS'],
    ['10', '2019', 'Telegram Network Monitoring Bot', 'Wrote a Telegram bot for real-time network monitoring at Goodrich Suites. Instant alerts on network events, dramatically faster incident response. First self-built automation project — got the itch from here.', 'Python,Telegram API,Automation,Network Monitoring'],
    ['11', '2023', 'SealInsure Entity IT Setup', 'Full IT setup for the newly launched SealInsure Indonesia entity — device provisioning, account setup, network configuration, and user onboarding. Operational from day one.', 'IT Infrastructure,Active Directory,Endpoint Management,ITSM'],
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

// ─── SKILLS ───────────────────────────────────────────────────────────────────
function populateSkills() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName('skills');
  if (!sheet) sheet = ss.insertSheet('skills');
  sheet.clearContents();

  const headers = ['skill', 'highlight'];
  const rows = [
    // Key competencies (highlight = true)
    ['IT Security & GRC',           'true'],
    ['ISO 27001 Lead Auditor',       'true'],
    ['AI & Automation',             'true'],
    ['Google Apps Script (GAS)',     'true'],
    ['Data Privacy (UU PDP)',        'true'],
    ['OJK Regulatory Compliance',    'true'],
    ['IT Risk Management',           'true'],
    // Technical (highlight = false)
    ['Custom GPT / LLM Integration', 'false'],
    ['Workflow Automation',          'false'],
    ['IT Service Management (ITSM)', 'false'],
    ['ITIL 4 Foundation',            'false'],
    ['Active Directory',             'false'],
    ['Network Administration',       'false'],
    ['Linux (Ubuntu, Proxmox)',      'false'],
    ['Endpoint Protection',          'false'],
    ['IT Asset Management',          'false'],
    ['Vendor Management',            'false'],
    ['IT Budgeting & CapEx',         'false'],
    ['Python / Scripting',           'false'],
    ['Google Workspace',             'false'],
    ['DNS, DHCP, TCP/IP',            'false'],
    ['Property Management Systems',  'false'],
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

// ─── TESTIMONIALS ─────────────────────────────────────────────────────────────
function populateTestimonials() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName('testimonials');
  if (!sheet) sheet = ss.insertSheet('testimonials');
  sheet.clearContents();

  const headers = ['name', 'title', 'relation', 'date', 'text'];
  const rows = [
    [
      'Faisal Rachmadianto',
      'Head of IT | Enterprise Technology Leader',
      'Faisal managed Eko directly',
      'May 2023',
      'I have the utmost pleasure recommending Kori for his exceptional work ethic and IT skills. He is a true gem who excels at working independently, taking risks, and consistently delivering IT services that surpass customer expectations. He is a joy to work with, always approachable and friendly, creating a positive work environment.'
    ],
    [
      'Armando Parhusip',
      'Head of IT | Head of Technology Operations',
      'Armando managed Eko directly',
      'Jul 2023',
      'I recommend Kori for his exceptional "can-do" attitude, unwavering eagerness to take on challenges, and consistent delivery of high-quality work. Throughout our collaboration, I was impressed by his ability to excel in a new role with little guidance, and his growth mindset made a positive impact. Keep up the excellent work, Kori!'
    ],
    [
      'Arbiter Gerhard Sarumaha',
      'General Manager',
      'Arbiter managed Eko directly',
      'Jun 2023',
      'Eko has a capability in IT field, finding immediate solutions and good at managing his works. His integrity is one of his strong characters. I was happy to work with him.'
    ],
    [
      'Andri Darmawan',
      'IT Corporate Manager | IT Infrastructure and Ops Manager',
      'Andri managed Eko directly',
      'Jun 2023',
      'Kori your nickname is friendly to me :), has good skill, innovative thinking, energik and give me the best solution related to technology.'
    ],
    [
      'Andrityo Hartoyo',
      'Passionate Hotelier',
      'Worked with Eko at different companies',
      'Aug 2023',
      'Eko Harianto always shown good work, he do his job well, eager to learn, fast learner, has good initiative and high motivation to gain success in his career.'
    ],
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

// ─── AGENT CONTEXT ────────────────────────────────────────────────────────────
function populateAgentContext() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName('agent_context');
  if (!sheet) sheet = ss.insertSheet('agent_context');
  sheet.clearContents();

  const lines = [
    ['Eko goes by "Kori" — he\'s open to IT leadership, GRC, cybersecurity, or IT automation/AI roles.'],
    ['Available for on-site, hybrid, or remote positions in Jakarta, Indonesia.'],
    ['Certifications: ISO 27001:2022 Lead Auditor, Certified DPO (Data Protection Officer), Qualified Chief Risk Officer (QCRO), ITIL 4 Foundation.'],
    ['Beyond governance, Kori actively builds internal tools: AI agents on the Quantum platform, Google Apps Script (GAS) automations for insurance ops (ticketing, mass notifications, compliance dashboards), and a Custom GPT for claims assessment.'],
    ['Career arc: Hospitality (2015–2018) → Hospitality IT Manager (2018–2020) → Enterprise IT at Shopee (2020–2023) → Head of IT at MoneeInsure digital insurance (2023–Present).'],
    ['At MoneeInsure (PT Asuransi Jiwa MoneeInsure), he wears two hats: IT security & compliance lead, and internal tooling/automation builder.'],
    ['He led the ISO 27001:2013 → 2022 transition audit with zero major findings and is the appointed DPO for UU PDP compliance.'],
    ['Technical depth: Linux (Ubuntu, Proxmox), Active Directory, network infrastructure (DNS, DHCP, firewall), Python scripting, Google Workspace, GAS, AI/LLM integration.'],
    ['He has 4,500+ LinkedIn followers. For salary/rate inquiries, direct to Eko personally via LinkedIn, email, or WhatsApp.'],
  ];

  sheet.getRange(1, 1, lines.length, 1).setValues(lines);
}
