from dotenv import load_dotenv
from pinecone import Pinecone, ServerlessSpec
from openai import OpenAI
import os
import time

# --- Load keys ---
load_dotenv()
pinecone_key = os.environ.get("PINECONE_API_KEY")
openai_key = os.environ.get("OPENAI_API_KEY")

if not pinecone_key or not openai_key:
    raise ValueError("Missing API key(s). Check your .env file.")

# --- Connect ---
pc = Pinecone(api_key=pinecone_key)
client = OpenAI(api_key=openai_key)

# --- Create index (only runs once, skips if it already exists) ---
index_name = "basic-search-demo"

if not pc.has_index(index_name):
    pc.create_index(
        name=index_name,
        dimension=1536,
        metric="cosine",
        spec=ServerlessSpec(cloud="aws", region="us-east-1")
    )
    print(f"Created index: {index_name}")
    time.sleep(5)  # give Pinecone a moment to finish setting up

index = pc.Index(index_name)

# --- Helper to turn text into a vector ---
def embed(text):
    return client.embeddings.create(
        model="text-embedding-3-small",
        input=text
    ).data[0].embedding

# --- Add sample data (only needs to run once) ---
data = [
    "Pinecone stores vectors and finds the closest matches to a query.",
    "Namespaces let you partition data within a single index.",
    "Metadata is extra info attached to a vector, like the original text or a source URL.",
    "RAG combines retrieval with generation to answer questions using your own data.",
]

vectors = [
    {"id": f"doc-{i}", "values": embed(text), "metadata": {"text": text}}
    for i, text in enumerate(data)
]

index.upsert(vectors=vectors)
print("Data uploaded.\n")

# --- Search ---
def search(question, top_k=3):
    query_vector = embed(question)
    results = index.query(vector=query_vector, top_k=top_k, include_metadata=True)
    print(f"Question: {question}\n")
    for match in results["matches"]:
        print(f"  score={match['score']:.3f} | {match['metadata']['text']}")
    print()

search("How do I organize data in Pinecone?")
search("What is RAG?")