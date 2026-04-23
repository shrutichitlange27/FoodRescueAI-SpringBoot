from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt

OUTPUT_FILE = "FoodProject_Thesis_Draft.docx"

SCREEN_SECTIONS = [
    {
        "title": "Home Page",
        "route": "/form/show",
        "file": "HomePage.html",
        "description": (
            "This is the landing page of the application. It introduces FoodRescueAI as a smart food waste management system "
            "and provides the main entry points for hotel, NGO, and administration users. The page also highlights the workflow, "
            "frequently asked questions, and a short impact summary to explain the purpose of the platform."
        ),
        "key_points": [
            "Hero section with project branding and call-to-action buttons.",
            "Separate portal cards for hotel and NGO registration and login.",
            "Workflow and FAQ sections that explain the donation flow.",
            "Administration login link for platform oversight."
        ],
    },
    {
        "title": "Hotel Registration Page",
        "route": "/form/hotel/register",
        "file": "RegForm.html",
        "description": (
            "This page allows hotels to register their organization details in the system. It collects the hotel name, permanent ID, "
            "mobile number, email, city, address, distance estimate, and food license proof. The registration data becomes the base profile "
            "used later for login, food listing, and admin verification."
        ),
        "key_points": [
            "Form fields for hotel identity and contact details.",
            "City list and address fields for matching nearby NGOs.",
            "File upload for food license proof.",
            "Validation messages for required data."
        ],
    },
    {
        "title": "Registration Success Page",
        "route": "/form/submit",
        "file": "Complete.html",
        "description": (
            "This confirmation page appears after successful hotel registration. It confirms that the hotel profile has been stored "
            "and provides a simple navigation link back to the home page. In the final thesis, this page can be used to show the end of the "
            "registration workflow."
        ),
        "key_points": [
            "Clear success message after form submission.",
            "Link back to the home page.",
            "Represents the end of the onboarding flow."
        ],
    },
    {
        "title": "Hotel Login Page",
        "route": "/form/hotel/login",
        "file": "UserLogin.html",
        "description": (
            "This page provides authentication for hotel users. Hotels log in using the permanent ID and mobile number that were created during registration. "
            "After login, the user is redirected to the hotel profile page where stored information and food management links are shown."
        ),
        "key_points": [
            "Login form for permanent ID and mobile number.",
            "Error display for invalid credentials.",
            "Link to hotel registration for new users."
        ],
    },
    {
        "title": "Hotel Profile Page",
        "route": "/form/hotel/details",
        "file": "UserProfile.html",
        "description": (
            "This page displays the stored hotel profile after a successful login. It shows the hotel name, permanent ID, contact number, address, city, and "
            "uploaded proof file path. The page also provides links to the food list manager and the hotel profile update form."
        ),
        "key_points": [
            "Read-only summary of hotel registration data.",
            "Link to add and view food inventory.",
            "Link to update profile details."
        ],
    },
    {
        "title": "Hotel Profile Update Page",
        "route": "/form/edit",
        "file": "Update.html",
        "description": (
            "This form is used to edit hotel profile details. Users can update their contact information, address, city, distance estimate, and license proof file. "
            "The page supports controlled profile maintenance without creating a new account."
        ),
        "key_points": [
            "Editable fields for hotel profile maintenance.",
            "Optional replacement of uploaded proof file.",
            "Back navigation to return to the previous view."
        ],
    },
    {
        "title": "Food List Management Page",
        "route": "/form/hotel/food",
        "file": "FoodList.html",
        "description": (
            "This is one of the main operational pages for hotel users. It allows the hotel to add new food items, edit existing items, "
            "track expiry status, review waste entries, and view claimed or delivered items. The page also shows NGO contact data once a donation is claimed."
        ),
        "key_points": [
            "Add food form with quantity, food type, dates, and photo upload.",
            "Inventory list for safe food items.",
            "Waste log for expired or unsafe food.",
            "Actions for edit, delete, claim handling, and delivery marking."
        ],
    },
    {
        "title": "Claim Response Page",
        "route": "/form/hotel/food/claims",
        "file": "FoodClaims.html",
        "description": (
            "This page is used to monitor how NGOs respond to hotel food listings. It shows claimed, rejected, and delivered records together with NGO contact information, "
            "hotel rejection reasons, and action buttons for delivery confirmation or rejection recording."
        ),
        "key_points": [
            "Unified table for all post-claim statuses.",
            "Delivery confirmation action for claimed items.",
            "Rejection form with reason field.",
            "Status pill labels for quick visual scanning."
        ],
    },
    {
        "title": "NGO Registration Page",
        "route": "/ngo/register",
        "file": "NGORegister.html",
        "description": (
            "This page collects the details of an NGO that wants to receive food donations. The form captures the NGO name, email, password, mobile number, city, "
            "preferred delivery distance, capacity, and address. These details are used later in matching and dashboard views."
        ),
        "key_points": [
            "NGO profile registration form.",
            "Preferred distance and meal capacity fields.",
            "City and address information for routing and matching."
        ],
    },
    {
        "title": "NGO Login Page",
        "route": "/ngo/login",
        "file": "NGOLogin.html",
        "description": (
            "This page authenticates registered NGO users using email and password. After login, NGOs are redirected to the dashboard, where they can review nearby donations, "
            "accept eligible food items, or reject items they cannot handle."
        ),
        "key_points": [
            "Login form with email and password fields.",
            "Error display for invalid or unverified users.",
            "Link to NGO registration."
        ],
    },
    {
        "title": "NGO Dashboard Page",
        "route": "/ngo/dashboard",
        "file": "NGODashboard.html",
        "description": (
            "This is the main working screen for NGO users. It shows the NGO profile, AI-suggested food donations, and additional nearby donations filtered by city and distance. "
            "Each food card contains hotel contact information and accept or reject actions."
        ),
        "key_points": [
            "NGO profile summary at the top.",
            "Suggested foods ranked by the rule engine.",
            "Nearby available donations list.",
            "Accept and reject buttons for each donation."
        ],
    },
    {
        "title": "NGO Profile Edit Page",
        "route": "/ngo/edit",
        "file": "NGOEdit.html",
        "description": (
            "This form allows NGOs to update their profile after registration. The NGO can change its name, contact number, city, address, preferred distance, capacity, "
            "and password while keeping the same registered email."
        ),
        "key_points": [
            "Editable NGO identity and contact fields.",
            "Capacity and distance updates for better matching.",
            "Save and return link to the dashboard."
        ],
    },
    {
        "title": "NGO Food Status Page",
        "route": "/ngo/accepted",
        "file": "NGOFoodStatus.html",
        "description": (
            "This page shows the food items that are already accepted, rejected, or delivered for a specific NGO. It acts as a history view, helping the NGO track completed and pending donation records in one table."
        ),
        "key_points": [
            "Historical list of donation decisions.",
            "Hotel contact and address information for each item.",
            "Status labels to distinguish accepted and delivered items."
        ],
    },
    {
        "title": "Admin Login Page",
        "route": "/set/login",
        "file": "AdminLogin.html",
        "description": (
            "This page provides access to the administration area of the system. The admin enters a username and password to enter the control panel where hotel records, NGO profiles, and food data can be managed."
        ),
        "key_points": [
            "Admin authentication form.",
            "Error display for invalid credentials.",
            "Access to the home page for quick navigation."
        ],
    },
    {
        "title": "Admin Dashboard Page",
        "route": "/set/dashboard",
        "file": "AdminDash.html",
        "description": (
            "This is the central administration dashboard. It contains two independent search panels, one for hotel records and another for NGO records. "
            "Administrators can open uploaded files, inspect hotel food records, and review NGO-related food lists from the same page."
        ),
        "key_points": [
            "Live search for hotels and NGOs.",
            "Table view of registration records.",
            "Links to hotel food inventory and NGO food history.",
            "Delete and logout actions for administration control."
        ],
    },
    {
        "title": "Admin Hotel Food View Page",
        "route": "/set/hotel-food/{permid}",
        "file": "AdminFoodView.html",
        "description": (
            "This page shows every food item listed by a selected hotel. It includes food details, expiry information, current status, photo links, and the name or contact number of the NGO if the item has been claimed."
        ),
        "key_points": [
            "Detailed hotel food inventory table.",
            "Photo preview link for each item.",
            "Link to edit individual food items."
        ],
    },
    {
        "title": "Admin Food Edit Page",
        "route": "/set/hotel-food/edit/{id}",
        "file": "AdminFoodEdit.html",
        "description": (
            "This edit form allows administrators to change food details for a hotel entry. The page highlights when a record is already claimed or delivered so that the workflow state remains protected while still allowing data corrections."
        ),
        "key_points": [
            "Editable food name, type, quantity, and date fields.",
            "Status lock notice for claimed or delivered records.",
            "Save and cancel controls."
        ],
    },
    {
        "title": "Admin NGO Food View Page",
        "route": "/set/ngo-food/{id}",
        "file": "AdminNGOFoods.html",
        "description": (
            "This page shows the foods associated with a selected NGO together with the hotel that posted each item. Administrators can delete a food entry or undo its status if correction is needed."
        ),
        "key_points": [
            "NGO-linked foods with hotel information.",
            "Delete action for food records.",
            "Undo action to restore a food listing."
        ],
    },
    {
        "title": "Mail Send Page",
        "route": "/send/mail-form/{name}/{email}",
        "file": "Mailsend.html",
        "description": (
            "This page is used to compose and send an email notification. The form keeps the recipient name and email prefilled, sets a default subject, and allows the administrator or workflow to send approval or completion messages."
        ),
        "key_points": [
            "Prefilled recipient name and email fields.",
            "Subject and message body area.",
            "Success or error feedback after sending."
        ],
    },
]


def add_run(paragraph, text, bold=False, italic=False, size=12):
    run = paragraph.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    return run


def add_heading(doc, text, level=1, center=False):
    paragraph = doc.add_paragraph()
    if center:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(text)
    run.bold = True
    run.font.name = "Times New Roman"
    run.font.size = Pt(18 if level == 1 else 14)
    paragraph.paragraph_format.space_after = Pt(6)
    return paragraph


def add_body_paragraph(doc, text, bold_prefix=None):
    paragraph = doc.add_paragraph()
    paragraph.style = doc.styles["Normal"]
    paragraph.paragraph_format.line_spacing = 1.5
    paragraph.paragraph_format.space_after = Pt(6)
    if bold_prefix and text.startswith(bold_prefix):
        add_run(paragraph, bold_prefix, bold=True)
        add_run(paragraph, text[len(bold_prefix):])
    else:
        add_run(paragraph, text)
    return paragraph


def add_bullets(doc, items):
    for item in items:
        paragraph = doc.add_paragraph(style="List Bullet")
        paragraph.paragraph_format.line_spacing = 1.5
        add_run(paragraph, item)


def page_break(doc):
    doc.add_page_break()


def add_cover_page(doc):
    add_heading(doc, "FoodRescueAI", level=1, center=True)
    add_body_paragraph(doc, "Thesis Draft for Food Waste Management and Donation Coordination System")
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_body_paragraph(doc, "Editable MS Word version prepared from the project source code.")
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_body_paragraph(doc, "Project Type: Spring Boot web application with hotel, NGO, and admin workflows.")
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_body_paragraph(doc, "Prepared for: Your College / Department / Guide")
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_body_paragraph(doc, "Prepared by: Your Name")
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_body_paragraph(doc, "Year: 2026")
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    page_break(doc)

    add_heading(doc, "Abstract")
    add_body_paragraph(doc, (
        "FoodRescueAI is a Spring Boot based platform designed to connect hotels that have surplus food with NGOs that can distribute it quickly. "
        "The system provides separate workflows for hotels, NGOs, and administrators. Hotels can register, log in, list food items, upload photos, "
        "and manage food status as items move from safe donation to claimed, delivered, rejected, or waste states. NGOs can register, log in, review "
        "nearby food listings, and accept or reject donations according to capacity and distance. Administrators supervise both sides through a dedicated "
        "dashboard that supports search, record inspection, and correction of food data."
    ))
    add_body_paragraph(doc, (
        "The application addresses the practical problem of food wastage by digitizing the donation process, preserving traceability, and reducing the time "
        "between food listing and pickup. The thesis draft below documents the project architecture, modules, and every HTML screen so that screenshots "
        "can be inserted later in Microsoft Word without changing the structure of the document."
    ))

    page_break(doc)

    add_heading(doc, "Acknowledgement")
    add_body_paragraph(doc, (
        "I would like to thank my guide, teachers, and classmates for their support during the development of this project. The FoodRescueAI application "
        "was built to explore a real-world social problem and turn it into a practical web system. I also acknowledge the open-source ecosystem, Spring Boot "
        "framework, and the HTML/CSS design approach used throughout the application."
    ))
    add_body_paragraph(doc, (
        "This thesis draft is intentionally written in an editable format so that names, institution details, and formatting preferences can be updated easily "
        "inside Microsoft Word."
    ))

    page_break(doc)

    add_heading(doc, "Table of Contents")
    add_body_paragraph(doc, "1. Introduction")
    add_body_paragraph(doc, "2. Problem Statement and Objectives")
    add_body_paragraph(doc, "3. System Overview and Architecture")
    add_body_paragraph(doc, "4. User Roles and Functional Modules")
    add_body_paragraph(doc, "5. Screen-wise Documentation with Screenshot Placeholders")
    add_body_paragraph(doc, "6. Conclusion and Future Scope")
    add_body_paragraph(doc, "This page can be replaced with an automatic Word table of contents if needed.")

    page_break(doc)

    add_heading(doc, "Chapter 1: Introduction")
    add_body_paragraph(doc, (
        "Food waste is a recurring operational and social problem in food service businesses. At the same time, many NGOs need a reliable and fast way to "
        "receive surplus food before it becomes unsafe. FoodRescueAI bridges this gap with a web-based workflow that lets hotels list available food and "
        "NGOs claim those donations through an organized digital process."
    ))
    add_body_paragraph(doc, (
        "The project uses Spring Boot for the backend, Thymeleaf templates for the user interface, and a MySQL-backed data model for persistent storage. "
        "The application has been structured around three roles: hotel user, NGO user, and administrator. Each role sees a different set of pages and actions."
    ))
    add_body_paragraph(doc, (
        "The rest of this document explains the project in thesis style and then documents every HTML page individually so the final submission can include "
        "the corresponding screenshots and captions."
    ))

    page_break(doc)

    add_heading(doc, "Chapter 2: Problem Statement and Objectives")
    add_body_paragraph(doc, (
        "The main problem addressed by this project is the delay and confusion that can happen when food donation is managed manually. Phone calls, informal "
        "messages, and untracked claims make it difficult to know which NGO received which food item and what the final delivery status is."
    ))
    add_body_paragraph(doc, (
        "The primary objective of FoodRescueAI is to create a single digital workflow for donation coordination. The platform stores hotel and NGO profiles, "
        "matches food items with nearby NGOs, tracks claim outcomes, and keeps an auditable history for both the hotel and the administrator."
    ))
    add_bullets(doc, [
        "Reduce food wastage by enabling fast donation listing.",
        "Help NGOs discover food donations that fit their city and distance limits.",
        "Maintain clear status tracking for claimed, rejected, and delivered items.",
        "Allow administrators to review and correct records when necessary."
    ])

    page_break(doc)

    add_heading(doc, "Chapter 3: System Overview and Architecture")
    add_body_paragraph(doc, (
        "The system follows a role-based web architecture. The home page acts as the landing layer, after which hotel users, NGO users, and admins move into "
        "their respective authentication pages. Once authenticated, hotels manage food listings, NGOs manage claim decisions, and admins manage both user groups "
        "and supporting records."
    ))
    add_body_paragraph(doc, (
        "The backend controllers separate the application logically into form management, food management, NGO operations, admin oversight, and email communication. "
        "This keeps the flow readable and supports a thesis narrative that is easy to explain in chapters."
    ))
    add_bullets(doc, [
        "FormController handles hotel registration, login, profile view, and profile updates.",
        "FoodController manages food creation, edits, deletion, claims, rejection, and delivery updates.",
        "NGOController manages NGO registration, dashboard matching, claim actions, and profile editing.",
        "AdminController manages dashboards, search, editing, and admin inspection pages.",
        "MailController provides a simple notification email page."
    ])

    page_break(doc)

    add_heading(doc, "Chapter 4: User Roles and Functional Modules")
    add_body_paragraph(doc, (
        "Hotel users are responsible for publishing surplus food, managing food freshness, and confirming final delivery or rejection details. NGO users are responsible "
        "for reviewing nearby items and deciding whether they can accept them. Administrators oversee the system and ensure data quality."
    ))
    add_bullets(doc, [
        "Hotel module: registration, login, profile, food inventory, claims, and status updates.",
        "NGO module: registration, login, dashboard matching, accept/reject workflow, and history.",
        "Admin module: dashboard search, hotel food review, NGO food review, edits, and cleanup.",
        "Messaging module: email notification form for communication and confirmations."
    ])
    add_body_paragraph(doc, (
        "This role separation is useful in a thesis because it explains why the interface contains many HTML pages and how each page serves a specific audience."
    ))

    page_break(doc)

    add_heading(doc, "Chapter 5: Data and Workflow Summary")
    add_body_paragraph(doc, (
        "Hotel and NGO profiles are stored as structured form data. Food records keep key fields such as quantity, food type, freshness dates, photo filename, status, "
        "and the references needed to connect the item to a hotel or NGO. Status values are used throughout the application to decide which items appear on each screen."
    ))
    add_body_paragraph(doc, (
        "The workflow starts with registration, continues through login and dashboard interaction, and ends with a delivery or rejection history. The admin layer "
        "can review stored data at any time. This makes the project suitable for a thesis because the logic is simple to explain but still meaningful in practice."
    ))

    page_break(doc)

    add_heading(doc, "Screen-wise Documentation and Screenshot Placeholders")
    add_body_paragraph(doc, (
        "The following pages document each HTML screen in the project. A screenshot placeholder is included in every section so you can paste the actual "
        "application screenshot into Microsoft Word later."
    ))


def add_screen_page(doc, screen):
    page_break(doc)
    add_heading(doc, screen["title"])
    add_body_paragraph(doc, f"Route: {screen['route']}")
    add_body_paragraph(doc, f"Template File: {screen['file']}")
    add_body_paragraph(doc, screen["description"])
    add_body_paragraph(doc, "Screenshot placeholder: insert the full-page screenshot of this screen here.")
    add_body_paragraph(doc, "Key elements on this screen:")
    add_bullets(doc, screen["key_points"])


def add_conclusion_pages(doc):
    page_break(doc)
    add_heading(doc, "Conclusion")
    add_body_paragraph(doc, (
        "FoodRescueAI demonstrates how a focused Spring Boot application can support an important community process. The project combines web forms, role-based "
        "dashboards, food status tracking, and admin oversight into one platform that is practical to explain in a thesis and useful in real deployment."
    ))
    add_body_paragraph(doc, (
        "The editable Word draft created here is intentionally structured so that the user can insert screenshots, update names and institute details, and "
        "refine any chapter before final submission."
    ))

    page_break(doc)
    add_heading(doc, "Future Scope")
    add_bullets(doc, [
        "Add automatic map-based distance calculation for better matching.",
        "Integrate a notification system for SMS or push alerts.",
        "Add analytics dashboards for donation volume and delivery time.",
        "Introduce role-based access control with stronger authentication.",
        "Add a PDF export option for donation reports and monthly summaries."
    ])

    page_break(doc)
    add_heading(doc, "References")
    add_bullets(doc, [
        "Spring Boot official documentation.",
        "Thymeleaf official documentation.",
        "MySQL documentation.",
        "Java 17 language documentation.",
        "Project source code and HTML templates from FoodRescueAI."
    ])


def main():
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.space_after = Pt(6)

    add_cover_page(doc)

    for screen in SCREEN_SECTIONS:
        add_screen_page(doc, screen)

    add_conclusion_pages(doc)

    doc.save(OUTPUT_FILE)
    print(OUTPUT_FILE)


if __name__ == "__main__":
    main()
