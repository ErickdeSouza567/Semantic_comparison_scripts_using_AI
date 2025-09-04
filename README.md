A script was developed to solve a real problem in my organization. We receive PDFs containing a large number of jobs that need to be inserted into the database, but we faced several challenges:

Many jobs already exist in the database.

The jobs need to be linked to existing categories and competencies in the database.

We needed a solution that works for a large number of job areas.

To address this, I developed a data pipeline that extracts jobs from PDFs and transforms them into a CSV file. I then used an offline AI model with embeddings to process all areas and contexts, classifying each job into the most appropriate categories based on semantic similarity.

Additionally, I created a script that analyzes each job with multiple competencies, establishes the correct relationships, and performs updates and inserts to ensure all data is accurately reflected in the database.

As a result, the pipeline processed 517 jobs using multithreading to speed up the process, classifying them into 35 categories and linking them with over 20,000 competencies.
