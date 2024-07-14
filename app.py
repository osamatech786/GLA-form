import streamlit as st
from streamlit_drawable_canvas import st_canvas
import shutil
from openpyxl import load_workbook
import re
from PIL import Image as PILImage
from openpyxl.drawing.image import Image as XLImage
from datetime import datetime, date
import smtplib
from email.message import EmailMessage
from datetime import date, timedelta

def app():
    st.set_page_config(
        page_title="Online Form",
        page_icon="📝",
        layout="wide",
        initial_sidebar_state="collapsed",
    )

    # Custom CSS to set a light background color
    st.markdown("""
        <style>
        body {
            background-color: #f0f0f0; /* Light grey background */
        }
        </style>
        """,
                unsafe_allow_html=True)

    st.image('header/header-GLA.png', use_column_width=True)

    files=list()

    st.title('Welcome')
    st.subheader('Please fill out the following details:')

    # Form Inputs
    first_name = st.text_input('First Name')
    middle_name = st.text_input('Middle Name')
    family_name = st.text_input('Family Name')

    date_of_birth = st.date_input(
    label="Select a date",
    value=datetime(2000, 1, 1),  # Default date
    min_value=date(1900, 1, 1),  # Minimum selectable date
    max_value=date(2025, 12, 31),  # Maximum selectable date
    key="date_input_widget",  # Unique key for the widget
    help="Choose a date"  # Tooltip text
)

    # st.header('Eligibility Check')

    # st.text("""
    #     Evidence CANNOT be accepted that has been entered at a later date than Actual End Date of the start aim.
    #     Evidence must be present for ALL 4 (EO1,2,3,4) of the below eligibility checks. Original documentation must have been witnessed by the Provider and preferably copies made as evidence in case of future audits.
    #     For list of ALL acceptable supporting documents check 'Start-Eligibility Evidence list'
    #     """)

    # st.text("""
    #     UK, EEA Nationals and Non-EEA Nationals

    #     EEA Countries (as of 27/01/2021): 
    #     Austria, Belgium, Bulgaria, Croatia, Republic of Cyprus, Czech Republic, Denmark, Estonia, Finland, France, Germany, Greece, Hungary, Ireland, Italy, Latvia, Lithuania, Luxembourg, Malta, Netherlands, Poland, Portugal, Romania, Slovakia, Slovenia, Spain, Sweden, Iceland, Liechtenstein, Norway.

    #     Switzerland is not an EU or EEA member but is part of the single market. This means Swiss nationals have the same rights to live and work in the UK as other EEA nationals.

    #     “Irish citizens in the UK hold a unique status under each country’s national law. You do not need permission to enter or remain in the UK, including a visa, any form of residence permit or employment permit”. Quote taken from below link:
    #     https://www.gov.uk/government/publications/common-travel-area-guidance/common-travel-area-guidance

    #     Non-EEA nationals who hold leave to enter or leave to remain with a permission to work (including status under the EUSS where they are an eligible family member of an EEA national) are also eligible for ESF support whilst in the UK.
    #     """)

    st.header('E01: Right to Live and Work in the UK')

    hold_settled_status, hold_pre_settled_status, hold_leave_to_remain = '', '', ''

    # Create a radio button for the Yes/No question
    british_or_not = st.radio(
        'Are you a UK OR Irish National OR European Economic Area (EEA) National?',
        ('Yes', 'No')
    )

    if british_or_not == 'Yes':
        nationality = st.text_input('Nationality')
        options = [
            'Full UK Passport',
            'Full EU Member Passport (must be in date - usually 10 years)',
            'National Identity Card (EU)'
        ]
        selected_option_nationality = st.radio("Select the type of document:", options)
        full_uk_passport, full_eu_passport, national_identity_card = '', '', ''

        if selected_option_nationality == options[0]:
            full_uk_passport, full_eu_passport, national_identity_card = 'X', '', ''
            st.text('Please upload a copy of your Full UK Passport')
            uploaded_file = st.file_uploader("Upload Full UK Passport", type=['docx', 'pdf', 'jpg', 'jpeg', 'png'])
            if uploaded_file is not None:
                files.append(uploaded_file)

        elif selected_option_nationality == options[1]:
            full_uk_passport, full_eu_passport, national_identity_card = '', 'X', ''
            st.text('Please upload a copy of your Full EU Member Passport')
            uploaded_file = st.file_uploader("Upload Full EU Member Passport", type=['docx', 'pdf', 'jpg', 'jpeg', 'png'])
            if uploaded_file is not None:
                files.append(uploaded_file)

        elif selected_option_nationality == options[2]:
            full_uk_passport, full_eu_passport, national_identity_card = '', '', 'X'
            st.text('Please upload a copy of your National Identity Card (EU)')
            uploaded_file = st.file_uploader("Upload National Identity Card (EU)", type=['docx', 'pdf', 'jpg', 'jpeg', 'png'])
            if uploaded_file is not None:
                files.append(uploaded_file)

        if selected_option_nationality in [options[1], options[2]]:
            st.text(
                'In order to be eligible for ESF funding, EEA Nationals must meet one of the following conditions'
            )
            conditions = [
                'a. Hold settled status granted under the EU Settlement Scheme (EUSS)',
                'b. Hold pre-settled status granted under the European Union Settlement Scheme (EUSS)',
                'c. Hold leave to remain with permission to work granted under the new Points Based Immigration System'
            ]
            settled_status = st.radio("Select your status:", conditions)

            if settled_status == conditions[0]:
                hold_settled_status, hold_pre_settled_status, hold_leave_to_remain = 'X', '', ''
                st.text('Please upload proof of your settled status')
                uploaded_file = st.file_uploader("Upload Proof of Settled Status", type=['docx', 'pdf', 'jpg', 'jpeg', 'png'])
                if uploaded_file is not None:
                    files.append(uploaded_file)

            elif settled_status == conditions[1]:
                hold_settled_status, hold_pre_settled_status, hold_leave_to_remain = '', 'X', ''
                st.text('Please upload proof of your pre-settled status')
                uploaded_file = st.file_uploader("Upload Proof of Pre-Settled Status", type=['docx', 'pdf', 'jpg', 'jpeg', 'png'])
                if uploaded_file is not None:
                    files.append(uploaded_file)

            elif settled_status == conditions[2]:
                hold_settled_status, hold_pre_settled_status, hold_leave_to_remain = '', '', 'X'
                st.text('Please upload proof of your leave to remain status')
                uploaded_file = st.file_uploader("Upload Proof of Leave to Remain Status", type=['docx', 'pdf', 'jpg', 'jpeg', 'png'])
                if uploaded_file is not None:
                    files.append(uploaded_file)

    else:
        not_nationality = st.text_input('Nationality ')
        passport_non_eu_checked = st.checkbox(
            'Passport from non-EU member state (must be in date) AND any of the below a, b, or c'
        )
        if passport_non_eu_checked:
            passport_non_eu = 'X'
            st.text('Please upload a copy of your non-EU Passport')
            uploaded_file = st.file_uploader("Upload Non-EU Passport", type=['docx', 'pdf', 'jpg', 'jpeg', 'png'])
            if uploaded_file is not None:
                files.append(uploaded_file)
        else:
            passport_non_eu = ''

        document_options = [
            "a. Letter from the UK Immigration and Nationality Directorate granting indefinite leave to remain (settled status)",
            "b. Passport either endorsed 'indefinite leave to remain' – (settled status) or includes work or residency permits or visa stamps (unexpired) and all related conditions met; add details below",
            "c. Some non-EEA nationals have an Identity Card (Biometric Permit) issued by the Home Office in place of a visa, confirming the participant’s right to stay, work or study in the UK – these cards are acceptable"
        ]

        document_type = st.radio("Select the type of document:", document_options)
        letter_uk_immigration, passport_endorsed, identity_card = '', '', ''

        if document_type == document_options[0]:
            letter_uk_immigration, passport_endorsed, identity_card = 'X', '', ''
            st.text('Please upload your Letter from the UK Immigration and Nationality Directorate')
            uploaded_file = st.file_uploader("Upload Letter from UK Immigration and Nationality Directorate", type=['docx', 'pdf', 'jpg', 'jpeg', 'png'])
            if uploaded_file is not None:
                files.append(uploaded_file)

        elif document_type == document_options[1]:
            letter_uk_immigration, passport_endorsed, identity_card = '', 'X', ''
            st.text('Please upload your endorsed passport')
            uploaded_file = st.file_uploader("Upload Endorsed Passport", type=['docx', 'pdf', 'jpg', 'jpeg', 'png'])
            if uploaded_file is not None:
                files.append(uploaded_file)

        elif document_type == document_options[2]:
            letter_uk_immigration, passport_endorsed, identity_card = '', '', 'X'
            st.text('Please upload your Identity Card (Biometric Permit)')
            uploaded_file = st.file_uploader("Upload Identity Card (Biometric Permit)", type=['docx', 'pdf', 'jpg', 'jpeg', 'png'])
            if uploaded_file is not None:
                files.append(uploaded_file)

        country_of_issue = st.text_input('Country of issue')
        id_document_reference_number = st.text_input('ID Document Reference Number')

        e01_date_of_issue = st.date_input(
            label="Date of Issue",
            value=datetime(2000, 1, 1),  # Default date
            min_value=date(1900, 1, 1),  # Minimum selectable date
            max_value=date(2025, 12, 31),  # Maximum selectable date
            help="Choose a date"  # Tooltip text
        )

        e01_date_of_expiry = st.date_input(
            label="Date of Expiry",
            value=datetime(2000, 1, 1),  # Default date
            min_value=date(1900, 1, 1),  # Minimum selectable date
            max_value=date(2025, 12, 31),  # Maximum selectable date
            help="Choose a date"  # Tooltip text
        )

        e01_additional_notes = st.text_area('Additional Notes',
                                            'Use this space for additional notes where relevant (type of Visa, restrictions, expiry etc.)')
        

    
    st.header('E02: Proof of Age (* all documents must be in date and if a letter is used, it must be within the last 3 months)')



    full_passport_eu = add_checkbox_with_upload('Full Passport (EU Member State)', 'full_passport_eu')
    national_id_card_eu = add_checkbox_with_upload('National ID Card (EU)', 'national_id_card_eu')
    firearms_certificate = add_checkbox_with_upload('Firearms Certificate/Shotgun Licence', 'firearms_certificate')
    birth_adoption_certificate = add_checkbox_with_upload('Birth/Adoption Certificate', 'birth_adoption_certificate')
    e02_drivers_license = add_checkbox_with_upload('Drivers Licence (photo card)', 'e02_drivers_license')
    edu_institution_letter = add_checkbox_with_upload('Letter from Educational Institution* (showing DOB)', 'edu_institution_letter')
    e02_employment_contract = add_checkbox_with_upload('Employment Contract/Pay Slip (showing DOB)', 'e02_employment_contract')
    state_benefits_letter = add_checkbox_with_upload('State Benefits Letter* (showing DOB)', 'state_benefits_letter')
    pension_statement = add_checkbox_with_upload('Pension Statement* (showing DOB)', 'pension_statement')
    northern_ireland_voters_card = add_checkbox_with_upload('Northern Ireland voters card', 'northern_ireland_voters_card')

    e02_other_evidence_text = st.text_input('Other Evidence: Please state type')

    # Validation for the last 3 months
    current_date = date.today()
    three_months_ago = current_date - timedelta(days=90)

    e02_date_of_issue = st.date_input(
        label="Date of Issue of evidence",
        value=date.today(),  # Default date
        min_value=date(1900, 1, 1),  # Minimum selectable date
        max_value=date(2025, 12, 31),  # Maximum selectable date
        help="Choose a date"  # Tooltip text
    )

    # Check if the selected date is within the last three months
    if e02_date_of_issue < three_months_ago:
        st.warning("The date of issue is not within the last 3 months. Please select a valid date.")
        st.stop()
    st.success("The date of issue is within the last 3 months.")
    

    st.header('E03: Proof of Residence (must show the address recorded on ILP) *within the last 3 months')

    e03_drivers_license = add_checkbox_with_upload('Drivers Licence (photo card)', 'e03_drivers_license')
    bank_statement = add_checkbox_with_upload('Bank Statement *', 'bank_statement')
    e03_pension_statement = add_checkbox_with_upload('Pension Statement*', 'e03_pension_statement')
    mortgage_statement = add_checkbox_with_upload('Mortgage Statement*', 'mortgage_statement')
    utility_bill = add_checkbox_with_upload('Utility Bill* (excluding mobile phone)', 'utility_bill')
    council_tax_statement = add_checkbox_with_upload('Council Tax annual statement or monthly bill*', 'council_tax_statement')
    electoral_role_evidence = add_checkbox_with_upload('Electoral Role registration evidence*', 'electoral_role_evidence')
    homeowner_letter = add_checkbox_with_upload('Letter/confirmation from homeowner (family/lodging)', 'homeowner_letter')

    e03_other_evidence_text = st.text_input('Other Evidence: Please state type', key='e03_other_evidence')

    # Validation for the last 3 months
    e03_date_of_issue = st.date_input(
        label="Date of Issue evidence",
        value=date.today(),  # Default date
        min_value=date(1900, 1, 1),  # Minimum selectable date
        max_value=date(2025, 12, 31),  # Maximum selectable date
        help="Choose a date",  # Tooltip text
        key='e03_date_of_issue'
    )

    # Check if the selected date is within the last three months
    if e03_date_of_issue < three_months_ago:
        st.warning("The date of issue is not within the last 3 months. Please select a valid date.")
        st.stop()
    st.success("The date of issue is within the last 3 months.")

    st.header('E04: Employment Status (please select one option from below and take a copy)')

    latest_payslip = '-'
    e04_employment_contract = '-'
    confirmation_from_employer = '-'
    redundancy_notice = '-'
    sa302_declaration = '-'
    ni_contributions = '-'
    business_records = '-'
    companies_house_records = '-'
    other_evidence_employed = '-'
    unemployed = '-'

    main_options = [
        'a. Latest Payslip (maximum 3 months prior to start date)',
        'b. Employment Contract',
        'c. Confirmation from the employer that the Participant is currently employed by them which must detail: Participant full name, contracted hours, start date AND date of birth or NINO',
        'd. Redundancy consultation or notice (general notice to group of staff or individual notifications) At risk of Redundancy only',
        'e. Self-employed',
        'f. Other evidence as listed in the \'Start-Eligibility Evidence list\' under Employed section - State below',
        'g. Unemployed (complete the Employment section in ILP form)'
    ]

    selected_main_option = st.radio("Select an employment status or document:", main_options, key="e04_main_option")

    if selected_main_option == main_options[0]:
        latest_payslip = handle_file_upload('Latest Payslip (maximum 3 months prior to start date)', 'latest_payslip')
    elif selected_main_option == main_options[1]:
        e04_employment_contract = handle_file_upload('Employment Contract', 'e04_employment_contract')
    elif selected_main_option == main_options[2]:
        confirmation_from_employer = handle_file_upload('Confirmation from the employer', 'confirmation_from_employer')
    elif selected_main_option == main_options[3]:
        redundancy_notice = handle_file_upload('Redundancy consultation or notice', 'redundancy_notice')
    elif selected_main_option == main_options[4]:
        self_employed_options = [
            "HMRC 'SA302' self-assessment tax declaration, with acknowledgement of receipt (within last 12 months)",
            'Records to show actual payment of Class 2 National Insurance Contributions (within last 12 months)',
            'Business records in the name of the business - evidence that a business has been established and is active / operating (within last 12 months)',
            'If registered as a Limited company: Companies House records / listed as Company Director (within last 12 months)'
        ]
        selected_self_employed_option = st.radio("Select self-employed evidence:", self_employed_options, key="e04_self_employed_option")
        if selected_self_employed_option == self_employed_options[0]:
            sa302_declaration = handle_file_upload("HMRC 'SA302' self-assessment tax declaration", 'sa302_declaration')
        elif selected_self_employed_option == self_employed_options[1]:
            ni_contributions = handle_file_upload('Records of Class 2 National Insurance Contributions', 'ni_contributions')
        elif selected_self_employed_option == self_employed_options[2]:
            business_records = handle_file_upload('Business records', 'business_records')
        elif selected_self_employed_option == self_employed_options[3]:
            companies_house_records = handle_file_upload('Companies House records', 'companies_house_records')
    elif selected_main_option == main_options[5]:
        other_evidence_employed = handle_file_upload("Other evidence as listed in the 'Start-Eligibility Evidence list'", 'other_evidence_employed')
    elif selected_main_option == main_options[6]:
        unemployed = handle_file_upload('Unemployed (complete the Employment section in ILP form)', 'unemployed')

    # Validation for the date of issue
    current_date = date.today()
    three_months_ago = current_date - timedelta(days=90)

    e04_date_of_issue = st.date_input(
        label="Date of Issue of evidence",
        value=date.today(),  # Default date
        min_value=date(1900, 1, 1),  # Minimum selectable date
        max_value=date(2025, 12, 31),  # Maximum selectable date
        help="Choose a date",  # Tooltip text
        key='e04_date_of_issue'
    )

    if e04_date_of_issue < three_months_ago:
        st.warning("The date of issue is not within the last 3 months. Please select a valid date.")
        st.stop()
    st.success("The date of issue is within the last 3 months.")

    st.header('Initial Assessment')
  
    qualification_or_training = st.radio(
    'Are you currently undertaking a qualification or training?',
    ['No', 'Yes'])

    if qualification_or_training=='Yes':
        qualification_or_training_y, qualification_or_training_n = 'Y', '-'
        st.subheader('Details of Qualification or Training')

        course_details = st.text_area('Course Details',
                                      'Enter details of the course')
        funding_details = st.text_area(
            'Funding Details', 'Enter details of how the course is funded')
    else:
        qualification_or_training_y, qualification_or_training_n = '-', 'N'
        course_details, funding_details = '', ''
        st.write(
            'You answered "No" to currently undertaking a qualification or training.'
        )

    st.header('Evidenced Qualification Levels')


    st.subheader('Participant self declaration of highest qualification level')
    participant_options = [
        'Below Level 1', 'Level 1', 'Level 2', 'Level 3', 'Level 4',
        'Level 5 and above', 'No Qualifications'
    ]


    participant_declaration = st.radio('', participant_options)


    p58 = '-'
    p59 = '-'
    p60 = '-'
    p61 = '-'
    p62 = '-'
    p63 = '-'
    p64 = '-'


    if participant_declaration == participant_options[0]:
        p58 = 'X'
    elif participant_declaration == participant_options[1]:
        p59 = 'X'
    elif participant_declaration == participant_options[2]:
        p60 = 'X'
    elif participant_declaration == participant_options[3]:
        p61 = 'X'
    elif participant_declaration == participant_options[4]:
        p62 = 'X'
    elif participant_declaration == participant_options[5]:
        p63 = 'X'
    elif participant_declaration == participant_options[6]:
        p64 = 'X'


    # st.subheader('Training Providers declaration')
    # training_provider_options = [
    #     'Below Level 1', 'Level 1', 'Level 2', 'Level 3', 'Below Level 4',
    #     'Level 5 and above', 'No Qualifications', 'No Personal Learning Record'
    # ]

    # training_provider_declaration = st.radio(
    #     'Please check the PLR and record information about prior attainment level to ensure correct recording of prior attainment, as well as ensuring no duplication of learning aims or units takes place.',
    #     training_provider_options)
    # p65 = '-'
    # p66 = '-'
    # p67 = '-'
    # p68 = '-'
    # p69 = '-'
    # p70 = '-'
    # p71 = '-'
    # p72 = '-'
    # justification='-'


    # if training_provider_declaration == training_provider_options[0]:
    #     p65 = 'X'
    # elif training_provider_declaration == training_provider_options[1]:
    #     p66 = 'X'
    # elif training_provider_declaration == training_provider_options[2]:
    #     p67 = 'X'
    # elif training_provider_declaration == training_provider_options[3]:
    #     p68 = 'X'
    # elif training_provider_declaration == training_provider_options[4]:
    #     p69 = 'X'
    # elif training_provider_declaration == training_provider_options[5]:
    #     p70 = 'X'
    # elif training_provider_declaration == training_provider_options[6]:
    #     p71 = 'X'
    # elif training_provider_declaration == training_provider_options[7]:
    #     p72 = 'X'

    # justification = st.text_area(
    #         'If there is a discrepancy between Participant self declaration and the PLR, please record justification for level to be reported'
    #     )

    # st.subheader('Does the participant have Basic Skills?')

    # english_options = ['none', 'Entry Level', 'Level 1', 'Level 2+']

    # english_skill = st.selectbox('English', english_options)

    # p74 = '-'
    # p75 = '-'
    # p76 = '-'
    # p77 = '-'

    # if english_skill == english_options[0]:
    #     p74 = 'X'
    # elif english_skill == english_options[1]:
    #     p75 = 'X'
    # elif english_skill == english_options[2]:
    #     p76 = 'X'
    # elif english_skill == english_options[3]:
    #     p77 = 'X'

    # maths_options = ['none', 'Entry Level', 'Level 1', 'Level 2+']

    # maths_skill = st.selectbox('Maths', maths_options)

    # p78 = '-'
    # p79 = '-'
    # p80 = '-'
    # p81 = '-'

    # if maths_skill == maths_options[0]:
    #     p78 = 'X'
    # elif maths_skill == maths_options[1]:
    #     p79 = 'X'
    # elif maths_skill == maths_options[2]:
    #     p80 = 'X'
    # elif maths_skill == maths_options[3]:
    #     p81 = 'X'

    # esol_options = ['none', 'Entry Level', 'Level 1', 'Level 2+']

    # esol_skill = st.selectbox('ESOL', esol_options)

    # p82 = '-'
    # p83 = '-'
    # p84 = '-'
    # p85 = '-'

    # if esol_skill == esol_options[0]:
    #     p82 = 'X'
    # elif esol_skill == esol_options[1]:
    #     p83 = 'X'
    # elif esol_skill == esol_options[2]:
    #     p84 = 'X'
    # elif esol_skill == esol_options[3]:
    #     p85 = 'X'

    # st.subheader('Basic Skills Initial Assessment')
    # st.text(
    #     "Initial Assessment Outcomes – record the levels achieved by the Participant"
    # )

    # maths_options = ['-', 'E1', 'E2', 'E3', '1', '2']

    # maths_level = st.selectbox('Maths Level', maths_options)

    # p86 = ''

    # if maths_level in maths_options[1:]:
    #     p86 = maths_level

    # english_options = ['-', 'E1', 'E2', 'E3', '1', '2']

    # english_level = st.selectbox('English Level', english_options)

    # p87 = ''

    # if english_level in english_options[1:]:
    #     p87 = english_level

    # st.subheader('Numeracy and Literacy Programmes')
    # completion_programmes = st.radio(
    #     'Will the Participant be completing relevant Numeracy and/or Literacy programmes within their learning plan?',
    #     ['Yes', 'No'])
    # p88 = '-'
    # p89 = '-'

    # if completion_programmes == 'Yes':
    #     p88 = 'Y'
    #     p89 = '-'
    # elif completion_programmes == 'No':
    #     p88 = '-'
    #     p89 = 'N'

    # st.subheader('Additional Learning Support')
    # additional_support = st.radio(
    #     'Does the Participant require additional learning and/or learner support?',
    #     ['Yes', 'No'])
    # p90 = '-'
    # p91 = '-'
    # support_details = '-'

    # if additional_support == 'Yes':
    #     p90 = 'Y'
    #     p91 = '-'
    #     support_details = st.text_area(
    #         'If answered \'Yes\' above, please detail how the participant will be supported'
    #     )
    # elif additional_support == 'No':
    #     p90 = '-'
    #     p91 = 'N'

    st.header('Current Skills, Experience, and IAG')

    st.subheader('Highest Level of Education at start')
    education_options = [
        'ISCED 0 - Lacking Foundation skills (below Primary Education)',
        'ISCED 1 - Primary Education',
        'ISCED 2 - GCSE D-G or 3-1/BTEC Level 1/Functional Skills Level 1',
        'ISCED 3 - GCSE A-C or 9-4/AS or A Level/NVQ or BTEC Level 2 or 3',
        'ISCED 4 - N/A',
        'ISCED 5 to 8 - BTEC Level 5 or NVQ Level 4, Foundation Degree, BA, MA or equivalent'
    ]

    # Change from selectbox to multiselect
    selected_levels = st.multiselect(
        'Select the highest level of education at start', education_options)

    # Initialize marks
    p93, p94, p95, p96, p97, p98 = '-', '-', '-', '-', '-', '-'

    # Mark selected options
    if education_options[0] in selected_levels:
        p93 = 'X'
    if education_options[1] in selected_levels:
        p94 = 'X'
    if education_options[2] in selected_levels:
        p95 = 'X'
    if education_options[3] in selected_levels:
        p96 = 'X'
    if education_options[4] in selected_levels:
        p97 = 'X'
    if education_options[5] in selected_levels:
        p98 = 'X'

    st.header('Other Information')


    st.subheader('Current Job Role and Day to Day Activities')
    job_role_activities = st.text_area(
        'What is your current job role and what are your day to day activities?'
    )


    st.subheader('Career Aspirations')
    career_aspirations = st.text_area('What are your career aspirations?')


    st.subheader('Training/Qualifications Needed')
    training_qualifications_needed = st.text_area(
        'What training/qualifications do you need to progress further in your career? (Planned and future training)'
    )


    st.subheader('Barriers to Achieving Career Aspirations')
    barriers_to_achieving_aspirations = st.text_area(
        'What are the barriers to achieving your career aspirations and goals?'
    )


    st.subheader('Courses/Programs Available')
    courses_programs_available = st.text_area(
        'What courses/programs/activity are available to you in order to meet your and your employer\'s needs?'
    )

    st.header('Induction Checklist')


    funded_by_mayor_of_london = st.checkbox(
        'This programme is funded by the Mayor of London')
    lls_completed = st.checkbox(
        'The London Learning Survey (LLS) has been completed and submitted')
    equality_diversity_policy = st.checkbox(
        'Equality and Diversity Policy/Procedure and point of contact')
    health_safety_policy = st.checkbox(
        'Health and Safety Policy/Procedure and point of contact')
    safeguarding_policy = st.checkbox(
        'Safeguarding Policy/Procedure and point of contact')
    prevent_policy = st.checkbox(
        'PREVENT and point of contact (including British Values)')
    disciplinary_policy = st.checkbox(
        'Disciplinary, Appeal and Grievance Policy/Procedures')
    plagiarism_policy = st.checkbox('Plagiarism, Cheating Policy/Procedure')
    terms_conditions = st.checkbox(
        'Terms and Conditions of Learning and programme content & programme delivery'
    )

    st.header('Declarations')


    # st.subheader('Provider Confirmation')
    st.text(
        'We hereby confirm that we have read, understood and agree with the contents of this document, and understand that the programme is funded by the Mayor of London.'
    )


    st.subheader('Participant Declaration')
    participant_declaration = st.text_area(
        'Participant Declaration',
        'I certify that I have provided all of the necessary information to confirm my eligibility for the Funded Provision.'
    )


    st.subheader('Participant Signature')

    st.text("Signature:")
    participant_signature = st_canvas(
        fill_color="rgba(255, 255, 255, 1)",  
        stroke_width=5,
        stroke_color="rgb(0, 0, 0)",  # Black stroke color
        background_color="white",  # White background color
        width=400,
        height=150,
        drawing_mode="freedraw",
        key="canvas",
    )

    date_signed = st.date_input(
    label="Date",
    value=datetime(2000, 1, 1),  # Default date
    min_value=date(1900, 1, 1),  # Minimum selectable date
    max_value=date(2025, 12, 31),  # Maximum selectable date
    help="Choose a date"  # Tooltip text
)

    submit_button = st.button('Submit')
    if submit_button:
        placeholder_values = {
            'p1': first_name,
            'p2': middle_name,
            'p3': family_name,
            'p4': date_of_birth,
            'p5': nationality,
            'p6': full_uk_passport,
            'p7': full_eu_passport,
            'p8': national_identity_card,
            'p9': hold_settled_status,
            'p10': hold_pre_settled_status,
            'p11': hold_leave_to_remain,
            'p12': not_nationality,
            'p13': passport_non_eu,
            'p14': letter_uk_immigration,
            'p15': passport_endorsed,
            'p16': identity_card,
            'p17': country_of_issue,
            'p18': id_document_reference_number,
            'p19': e01_date_of_issue,
            'p20': e01_date_of_expiry,
            'p21': e01_additional_notes,
            'p22': full_passport_eu,
            'p23': national_id_card_eu,
            'p24': firearms_certificate,
            'p25': birth_adoption_certificate,
            'p26': e02_drivers_license,
            'p27': edu_institution_letter,
            'p28': e02_employment_contract,
            'p29': state_benefits_letter,
            'p30': pension_statement,
            'p31': northern_ireland_voters_card,
            'p32': e02_other_evidence_text,
            'p33': e02_date_of_issue,
            'p34': e03_drivers_license,
            'p35': bank_statement,
            'p36': pension_statement,
            'p37': mortgage_statement,
            'p38': utility_bill,
            'p39': council_tax_statement,
            'p40': electoral_role_evidence,
            'p41': homeowner_letter,
            'p42': e03_date_of_issue,
            'p43': e03_other_evidence_text,
            'p44': latest_payslip,
            'p45': e04_employment_contract,
            'p46': confirmation_from_employer,
            'p47': redundancy_notice,
            'p48': sa302_declaration,
            'p49': ni_contributions,
            'p50': business_records,
            'p51': companies_house_records,
            'p52': other_evidence_employed,
            'p53': unemployed,
            'p54': e04_date_of_issue,
            'p55': qualification_or_training_y,
            'p56': qualification_or_training_n,
            'p57': course_details + ' ' + funding_details,
            'p58': p58,
            'p59': p59,
            'p60': p60,
            'p61': p61,
            'p62': p62,
            'p63': p63,
            'p64': p64,
            'p65': p65,
            'p66': p66,
            'p67': p67,
            'p68': p68,
            'p69': p69,
            'p70': p70,
            'p71': p71,
            'p72': p72,
            'p73': justification,
            'p74': p74,
            'p75': p75,
            'p76': p76,
            'p77': p77,
            'p78': p78,
            'p79': p79,
            'p80': p80,
            'p81': p81,
            'p82': p82,
            'p83': p83,
            'p84': p84,
            'p85': p85,
            'p86': p86,
            'p87': p87,
            'p88': p88,
            'p89': p89,
            'p90': p90,
            'p91': p91,
            'p92': support_details,
            'p93': p93,
            'p94': p94,
            'p95': p95,
            'p96': p96,
            'p97': p97,
            'p98': p98,
            'p99': job_role_activities,
            'p100': career_aspirations,
            'p101': training_qualifications_needed,
            'p102': barriers_to_achieving_aspirations,
            'p103': courses_programs_available,
            # 'p113': participant_signature,
            'p114': date_signed,
        }

        # mandatory fields validation
        exclude_fields = {'p1000', 'p1', 'p2', 'p3', 'p5', 'p7', 'p8', 'p10', 'p11', 'p12', 'p13', 'p15', 'p16', 'p17', 'p18', 'p32', 'p43', 'p73', 'p86', 'p87', 'p92', 'p99', 'p100', 'p101', 'p102', 'p103'}     # exclude fields
        mandatory_fields = [f'p{i}' for i in range(1, 151) if f'p{i}' not in exclude_fields]    # define mandatory fields
        missing_fields = validate_inputs(placeholder_values, mandatory_fields)  # get the list of missing mandatory inputs
        if missing_fields:
            st.warning(f"Please fill out all the fields.")
            st.text(missing_fields)
        else:        
            # Define input and output paths
            template_file = "ph gla.xlsx"
            modified_file = "Filled_GLA_AEB_start_forms.xlsx"

            if len(participant_signature.json_data['objects']) != 0:
                # Convert the drawing to a PIL image and save it
                signature_path = 'signature_image.png'
                signature_image = PILImage.fromarray(
                    participant_signature.image_data.astype('uint8'), 'RGBA')
                signature_image.save(signature_path)
                # st.success("Signature image saved!")

                # Multi Sheet Support
                sheet_names = ['Eligibility', 'ILR']

                replace_placeholders(template_file, modified_file,
                                    placeholder_values, signature_path, sheet_names)
                # st.success(f"Template modified and saved as {modified_file}")
                st.success('Form submitted successfully!')

                # Email
                # Sender email credentials
                sender_email = st.secrets["sender_email"]
                sender_password = st.secrets["sender_password"]
                
                receiver_email = sender_email
                # Email details
                subject = "GLA Form Submission Received"
                body = "GLA Form submitted. Please find the attached file."
                file_path = "Filled_GLA_AEB_start_forms.xlsx" 
                # Send email with attachment
                # send_email_with_attachment(sender_email, sender_password, receiver_email, subject, body, file_path)

            
            else:
                st.warning("Please draw your signature.")


#  Custome Functions 
# ####################################################################################################################################
def validate_inputs(inputs, mandatory_fields):
    """Check if all mandatory input fields are filled and return the list of missing fields."""
    missing_fields = []
    for key, value in inputs.items():
        if key in mandatory_fields and (value is None or value == ''):
            missing_fields.append(key)
    return missing_fields

def resize_image_to_fit_cell(image_path, max_width, max_height):
    with PILImage.open(image_path) as img:
        img.thumbnail((max_width, max_height), PILImage.Resampling.LANCZOS)
        return img


def replace_placeholders(template_file, modified_file, placeholder_values, signature_path, sheet_names):
    # Copy the template file to a new file
    shutil.copyfile(template_file, modified_file)

    # Load the new copied workbook
    wb = load_workbook(modified_file)

    for sheet_name in sheet_names:
        sheet = wb[sheet_name]

        # Replace placeholders with provided values or images
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    for placeholder, value in placeholder_values.items():
                        # Use regular expressions to find full placeholder word
                        pattern = re.compile(r'\b' + re.escape(placeholder) + r'\b')
                        cell.value = pattern.sub(str(value), cell.value)
                        if 'p113' in cell.value:
                            cell.value = cell.value.replace('p113', '')  
                            resized_image = resize_image_to_fit_cell(signature_path, 200, 55)
                            resized_image_path = 'resized_signature_image.png'
                            resized_image.save(resized_image_path)
                            img = XLImage(resized_image_path)
                            sheet.add_image(img, cell.coordinate)

    # Save the workbook
    wb.save(modified_file)

    # file download button
    with open(modified_file, 'rb') as f:
        file_contents = f.read()
        st.download_button(
            label="Download File",
            data=file_contents,
            file_name=modified_file,
            mime='application/octet-stream'
        )

def send_email_with_attachment(sender_email, sender_password, receiver_email, subject, body, file_path):
    """
    Send an email to yourself with an attachment.

    Args:
        sender_email (str): The sender's email address.
        sender_password (str): The sender's email password.
        receiver_email (str): The receiver's email address (same as sender's in this case).
        subject (str): The subject of the email.
        body (str): The body of the email.
        file_path (str): The path to the file to be attached.
    """
    msg = EmailMessage()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.set_content(body)

    # Read the file and attach it to the email
    with open(file_path, 'rb') as f:
        file_data = f.read()
        file_name = file_path.split('/')[-1]
        msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

    # Use the SMTP server for Microsoft email accounts
    with smtplib.SMTP('smtp.office365.com', 587) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)

# Function to add a checkbox with a file upload option
def add_checkbox_with_upload(label, key_prefix):
    checked = st.checkbox(label, key=f"{key_prefix}_checkbox")
    if checked:
        st.text(f'Please upload a copy of your {label}')
        uploaded_file = st.file_uploader(f"Upload {label}", type=['pdf', 'jpg', 'jpeg', 'png', 'docx'], key=f"{key_prefix}_uploader")
        if uploaded_file is not None:
            files.append(uploaded_file)
        return 'X'
    else:
        return '-'
    
def handle_file_upload(label, key_prefix):
    st.text(f'Please upload a copy of your {label}')
    uploaded_file = st.file_uploader(f"Upload {label}", type=['pdf', 'jpg', 'jpeg', 'png', 'docx'], key=key_prefix)
    if uploaded_file is not None:
        files.append(uploaded_file)
        return 'X'
    else:
        return '-'
    
if __name__ == '__main__':
    app()