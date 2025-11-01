# Gmail Fast Reader

## Goal

This project is a GMail add-on that allows the user to quickly find, flag, and 
summarize topics that are interesting for them.

## Workflow

There are three workflows that Gmail Fast Reader supports:

1. Passive workflow
   
   Gmail Fast Reader would go over the last day emails, find the emails 
   containing insteresting topics, summarize and label them, and create a new 
   daily email that contains the summary of the interesting topics of the day.

2. Interactive workflow

   A semi-interactive mode where the user would activate the Gmail Fast Reader 
   to go over the emails that arrived during the latest X hours or days, 
   summarize them, and create a new email that contains the summary of the 
   interesting topics of the emails selected.

3. Configuration

   An interactive mode where the user would configure parameters of the 
   Gmail Fast Reader.

## Filtering and summarization approach

Gmail Fast Reader will go over emails and fit them in two categories:

- I must do
- I must know

In every category the user will define a series of topics of interest, 
for example: In "I must do" they can specify "payments for school, tennis club
tournaments". In "I must know" they can specify things like "upcoming 
parent-teacher interviews at my daughter's school". Both categories support 
"other" flag that lets Gmail Fast Reader decide.

When going over emails, Gmail Fast Reader will perform three actions:

1. Determine whether the email fits one of the "interesting" categories.
2. If so, determine which one(s) and what is the key action or knowledge for 
   the user.
3. Try to infer the earliest date of action, event, or other time-bound 
   activity.

After all emails' scanning is done, Gmail Fast Reader will create a summary,
sorting the topics of interest by date. If the upcoming date is within 1 day 
(today or tomorrow), Gmail Fast Reader will mark it as "urgent".

Finally Gmail Fast Reader will send the summary to the user.

## Configuration parameters

Gmail Fast Reader would accept the following configuration parameters:

- Name

  A user-friendly name of their add-on instance, for example, Jeeves. 
  By default, the name is "Gmail Fast Reader".

- Topics

  Within the built-in categories, Gmail Fast Reader will allow the user to 
  enter a newline-separated list of topics of interest. In addition to that, it
  will allow the user to select the "Other" option.

- OpenAI key

  The user must specify their own OpenAI key for the add-on's OpenAI analysis
  backend to work.

- Time zome

  The user's time zone. The time zone will be used to determine the best time
  to trigger time-based tasks.

## Algorithm

Gmail Fast Reader will take every email and run them through an OpenAI's 
model that would find that email's relevance to one or more topics identified 
by the user.

If the relevance is found, Gmail Fast Reader would infer the key action or 
knowledge item and the earliest key date with the help of an OpenAI model.

