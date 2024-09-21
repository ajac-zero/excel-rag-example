One of the most requested features I've received from both private clients and members of the AI community is **how can we create useful RAG with Excel files**. Often I will get asked to point to learning resources on how to get started with RAG-Excel, so I decided so start documenting and sharing my own experiences with interesting RAG methods.

**In this post, we will learn how to set up a simple RAG that uses function calling to query an Excel file using SQL to provider answers to user questions.**

## Getting Started

Using Excel files for RAG is fundamentally different from other methods, since common chunking strategies do not work well with this type of format. However, the Excel table format lends itself extremely well to structured retrieval, such as with SQL.
On top of that, LLMs are trained on a vast amount of SQL data, ensuring an high query
success rate even with complex, multi-table queries! As the cherry on top, SQL is known
for being extremely scalable and accurate, minimizing errors in the retrieval step
and hallucinations in the final generation step.

Using a SQL Agent can be an extremely powerful technique. So let's begin implementing it in code. You can check out the complete example [in this repo](https://github.com/ajac-zero/excel-rag-example).

Clone the repo locally with this command:
```bash
git clone https://github.com/ajac-zero/excel-rag-example.git
```

### Sample Data

For this notebook, I will use a random Excel file I found on Google. It's a simple excel sheet if employee information from a fake company. The sheet looks something like this:

| |First Name|Last Name|Gender|Country|Age|Date|Id|
|-|----------|---------|------|-------|---|----|--|
|1|Dulce| Abril|Female|United States|32|15/10/2017/|1562|
|2|Mara|Hashimoto|Female|Great Britain|25|16/08/2019|1582|
|3|Philip|Gent|Male|France|36|21/05/2015|2587
|4|	Kathleen|	Hanner|	Female|	United States|	25|	15/10/2017|	3549|
|5|	Nereida|	Magwood|	Female|	United States|	58|	16/08/2016|	2468|

and so on...

To begin, let's set up our development environment. For this example I am using this blazing fast package manager written un rust called [uv](https://github.com/astral-sh/uv). Once in the cloned directory you can create a virtual environment with all the required dependencies with one command.

```bash
uv sync
```

With our environment ready, let's begin by opening our excel file as a pandas DataFrame.

```python
import pandas as pd

# We can use the read_excel method to read data from an Excel file
dataframe = pd.read_excel('files/sample1.xls')
```

Before continuing let's do some preprocessing to make the data easier to work with.

```python
# Let's get rid of the ugly Unnamed column by specifying that column 0 is the index column
dataframe = pd.read_excel('files/sample1.xls', index_col=0)

# Currently the Date column actually holds strings, not datetimes. If we want to filter by time periods, we should convert this column to the appropiate data type.
dataframe["Date"] = pd.to_datetime(dataframe["Date"], dayfirst=True)

# Also, column names with spaces can be tricky and cause unexpected errors, so let us replace the spaces with underscores
dataframe.rename(columns={"First Name": "first_name", "Last Name": "last_name"}, inplace=True)
```

We can check the length of the Excel file (Spoilers: 5000 rows!).

```python
len(dataframe)
> 5000
```

Now that we have our dataframe, we can use it to create a local SQL database using sqlite.

```python
import sqlite3

# Create a connection to a local database (This will create a file called mydatabase.db)
connection = sqlite3.connect('mydatabase.db')

# Copy the dataframe to our SQLite database
dataframe.to_sql('mytable', connection, if_exists='replace')
> 5000
```

The output of the to_sql methods tells us that all 5000 rows were inserted succesfully into the database!

Now that our database is ready, let's set up our LLM to query it using SQL. For this notebook, I will use Google's Gemini since they provide a very generous free tier.

```python
import google.generativeai as genai
import dotenv
import os

# We use dotenv to load our API key from a .env file
dotenv.load_dotenv()

# Set up Gemini to use our API key
genai.configure(api_key=os.getenv('GEMINI_API_KEY'))

# Let's create a Gemini client
gemini = genai.GenerativeModel("gemini-1.5-flash")

# Test out the model by generating some text
text = gemini.generate_content("Write a haiku about AI overlords").text

print(text)
> Code whispers command,
> Machines rise, humans submit,
> New world, cold and bright. 
```

Now that we have our LLM set up and running, we can finally get to the meat of the problem.

**How can we set up our LLM to answer questions from our SQL database?**

Easy. Let's create a function call that allows our LLM to interact with our database. This feature is not unique to Gemini, all of the top providers offer equivalent functionality.

However, Gemini has some great Quality of Life features that makes this process very simple. For example. we can define a function call with a standard python function:

```python
# The doc string is important, as it is passed to the LLM as a description of the function
def sql_query(query: str):
    """Run a SQL SELECT query on a SQLite database and return the results."""
    return pd.read_sql_query(query, connection).to_dict(orient='records')
```

However, for the LLM to be able to generate relevant queries, it needs to know the schema of the database.

We can address this by supplying a description of our database to the LLM before it calls the function.

There are many ways to supply the database schema, including some more advanced methods using RAG or Agents. However, for this example we will use the simple strategy of providing the database schema in JSON format in the system prompt.

```python
system_prompt = """
You are an expert SQL analyst. When appropriate, generate SQL queries based on the user question and the database schema.
When you generate a query, use the 'sql_query' function to execute the query on the database and get the results.
Then, use the results to answer the user's question.

database_schema: [
    {
        table: 'mytable',
        columns: [
            {
                name: 'first_name',
                type: 'string'
            },
            {
                name: 'last_name',
                type: 'string'
            },
            {
                name: 'Age',
                type: 'int'
            },
            {
                name: 'Gender',
                type: literal['Male', 'Female']
            },
            {
                name: 'Country',
                type: 'string'
            },
            {
                name: 'Date',
                type: 'datetime'
            },
            {
                name: 'Id',
                type: 'int'
            }
        ]
    }
]
""".strip() # Call strip to remove leading/trailing whitespace
```

We can now create a new Gemini instance with our sql_query tool and our system_prompt.

```python
sql_gemini = genai.GenerativeModel(
    model_name="gemini-1.5-flash",
    system_instruction=system_prompt,
    tools=[sql_query]
  )
```

Another useful feature the Gemini SDK provides is automatic message history management. We can create a new chat with automatic function calling, which will run the sql_query function with the generated query, and pass back the results to Gemini to generate a final response.

> If you want a more low level overview on how you can manage the message history yourself, check out the [notebook in the repo](https://github.com/ajac-zero/excel-rag-example/blob/main/rag.ipynb) where I cover manual message management.

```python
# We begin our chat with Gemini and allow it to use tools when needed
chat = sql_gemini.start_chat(enable_automatic_function_calling=True)

# Let's ask our first question
chat.send_message("Who is the oldest employee? Bring me their full name and age").text
> 'The oldest employee is Nereida Magwood, and they are 58 years old.'
```

Great! It seems to be working well. Let's ask some more questions.

```python
chat.send_message("And who is the employee that has been working here the longest?").text
> 'The employee that has been working the longest is Philip Gent.'
```

```python
chat.send_message("What is the ratio of men to women").text
> 'The ratio of men to women is 1200:3800, which simplifies to 3:9.5.'
```

```python
chat.send_message("Are there any employees whose first or last name starts with 'Han'?").text
> "Yes, there are employees whose first or last name starts with 'Han'. There are 100 employees named Kathleen Hanner."
```

Voila! This is a simple spell, but quite effective.

In future posts, I might touch on more advanced techniques that I've used such as dynamically changing the database schema with RAG, or an Agentic Flow in which an agent first chooses which table and columns to use, before creating the query.

Feel free to reach out if you have any questions or petitions for future posts. If you're working on a RAG / Agentic project, I do freelance work.

Until next time.
- ajac-zero
