from langchain_openai import ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import StrOutputParser
from langchain_community.llms import ollama

import streamlit as st
from dotenv import load_dotenv
import os

# Load environment variables from .env file
load_dotenv()

# Set the OpenAI API key if available
# os.environ["OPENAI_API_KEY"] = os.getenv("OPENAI_API_KEY")

# Enable LangChain tracing
# os.environ["LANGCHAIN_TRACING_V2"] = "true"

# Set the LangChain API key if available
# os.environ["LANGCHAIN_API_KEY"] = os.getenv("LANGCHAIN_API_KEY")


# Prompt template
chat_prompt = ChatPromptTemplate.from_messages([
    ("system", "You are a helpful assistant that answers questions based on the provided context."),
    ("human", "Question: {question}"),
])

# Streamlit app configuration
st.title("LLM Chatbot Demo With LangChain and Gemma 3 API")
input_text = st.text_input("Enter the question you want to ask:", "")

# OpenAI chat model
llm = ollama.Ollama(model="gemma3", temperature=0.7)
# Output parser
output_parser = StrOutputParser()
chain = chat_prompt | llm | output_parser

if input_text:
    with st.spinner("Generating response..."):
        try:
            # Run the chain with the input text
            response = chain.invoke({"question": input_text})
            st.success(f"Response: {response}")
        except Exception as e:
            st.error(f"Error: {str(e)}")