import pandas as pd

def load_excel(file):
    return pd.read_excel(file)

def save_to_excel(df, file_path):
    df.to_excel(file_path, index=False)

def generate_prompt(row):
    template = """
    As an expert email marketer at Cvent Company, your task is to craft a personalized email sequence for Mr. {contact_name}, who holds the position of {Title} and is a {persona} at the company named {company_name}, which operates in the {Industry} industry.

    The email sequence should consist of 1 emails and the email campaign is aimed at improving {Purpose_of_email_campaign}.

    The company {company_name} has {Parent_company} as its parent company, which is already a Cvent customer. Cvent has previously assisted {Parent_company} with {Relation_with_parent_company}. Mr. {contact_name} has {Contact_Engagement_data}.

    In each email, pay special attention to the client's role and responsibilities, industry challenges, and the purpose of the email. Stress on the point of existing relationship with parent's company and give special attention to it.

    Incorporate this information into the emails to demonstrate Cvent's expertise in providing software solutions for event planners and marketers, focusing on online event registration, venue selection, event management and marketing, virtual, hybrid, and onsite solutions, and attendee engagement.

    Your goal is to personalize the email sequence to address the client's specific needs and to effectively showcase the value of the software solutions offered by Cvent Company to the {Industry}.

    Please ensure that each email in the sequence is tailored to engage the client, address their challenges, and clearly communicate the benefits of the Cvent's software solutions as bulleted points where ever required in a compelling manner.

    Your response should be detailed, persuasive, and demonstrate a deep understanding of the client's industry and requirements. Be creative and flexible in crafting the emails, aiming to establish a strong rapport and drive conversion through personalized and expertly curated content.

    Include a dummy hyper link for calendar at the end of each email.

    The email sequence SHOULD CONSIST OF 1 EMAILS and the email campaign is aimed at improving {Purpose_of_email_campaign}.

    The output should be structured as follows:

    Email 1:

    Subject Line: [Concise subject line]
    Body: [Organized content with distinct points, challenges, benefits, and CTA]
    Signature: Praveen, Business Development Manager, Cvent
   """
    return template.format(
        contact_name=row['contact name'],
        Title=row['Title'],
        persona=row['persona'],
        company_name=row['company name'],
        Industry=row['Industry'],
        Parent_company=row['Parent company'],
        Relation_with_parent_company=row['Relation with parent company'],
        Contact_Engagement_data=row['Contact Engagement data'],
        Purpose_of_email_campaign=row['Purpose of email campaign']
    )
def generate_prompts(df):
    df['personalized_prompt'] = df.apply(generate_prompt, axis=1)
    return df
