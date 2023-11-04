import os
import openai

openai.api_key = "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"

def chatGPTApi(query):
    response = openai.ChatCompletion.create(
    model="gpt-3.5-turbo",
    messages=[
        {
        "role": "system",
        "content": "Generate only plain text paragraphs, each pragraph with a limit of around 50 words. You are allowed to include bullet points if required. You are allowed to use limited set of special characters like -, $, &, @, !, (, and )"
        },
        {
        "role": "user",
        "content": query
        }
    ],
    temperature=0.5,
    max_tokens=256,
    top_p=1,
    frequency_penalty=0,
    presence_penalty=0
    )
    # print(response)
    
    return response.choices[0].message.content

