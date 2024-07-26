import pandas as pd

def load_excel(file):
    return pd.read_excel(file)

def save_to_excel(df, file_path):
    df.to_excel(file_path, index=False)

def generate_prompt(row):
    template = """
    As an expert email marketer at Cvent Company, your task is to craft a personalized email sequence for Mr. {contact_name}, who holds the position of {Title} and is a {persona} at the company named {company_name}, which operates in the {Industry} software industry. The email sequence should consist of 3 emails, aimed at improving conversion for the campaign.

    The company {company_name} has {Parent_company} as its parent company, which is already a Cvent customer. Cvent has previously assisted {Parent_company} with {Relation_with_parent_company}. Mr. {contact_name} has {Contact_Engagement_data}. Each email should pay special attention to the client's role and responsibilities, industry challenges, and the purpose of the email. The emails should showcase Cvent's expertise in providing solutions for event planners and marketers, including online event registration, venue selection, event management and marketing, virtual, hybrid, and onsite solutions, and attendee engagement.

    Ensure that each email is personalized to address the client's specific needs and demonstrates the value of Cvent's software solutions. The content should be compelling, addressing the client's challenges, and include bulleted benefits where appropriate. Each email must include a CTA (Call to Action) at the end.

    The output should be structured as follows:

    Email 1:

    Subject Line: [Concise subject line]
    Body: [Organized content with distinct points, challenges, benefits, and CTA]
    Signature: [Full Name, Job Title, Company Name, Contact Information]

    Email 2:

    Subject Line: [Concise subject line]
    Body: [Organized content with distinct points, challenges, benefits, and CTA]
    Signature: [Full Name, Job Title, Company Name, Contact Information]

    Email 3:

    Subject Line: [Concise subject line]
    Body: [Organized content with distinct points, challenges, benefits, and CTA]
    Signature: [Full Name, Job Title, Company Name, Contact Information]

    Make sure that each email is tailored to engage the client, address their challenges, and clearly communicate the benefits of Cvent's software solutions. Be creative and flexible in crafting the emails to establish a strong rapport and drive conversion through personalized content.
    """
    return template.format(
        contact_name=row['contact name'],
        Title=row['Title'],
        persona=row['persona'],
        company_name=row['company name'],
        Industry=row['Industry'],
        Parent_company=row['Parent company'],
        Relation_with_parent_company=row['Relation with parent company'],
        Contact_Engagement_data=row['Contact Engagement data']
    )

def generate_prompts(df):
    df['personalized_prompt'] = df.apply(generate_prompt, axis=1)
    return df
