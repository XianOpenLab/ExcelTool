import openai
import json

openai.api_key = "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"
communication = ""
AI = ""
while True:
    human = input("Human:")
    communication += "Human:" + human + " \nAI:"

    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=communication,
        temperature=0.9,
        max_tokens=3000,
        top_p=1,
        frequency_penalty=0.0,
        presence_penalty=0.6,
        stop=[" Human:", " AI:"]
    )
    json_str = json.dumps(response)
    dic = json.loads(json_str)
    AI = dic["choices"][0]["text"]
    print("AI:" + AI)
    communication += AI + "\n"
    if len(dic["choices"]) > 1:
        print(response)
