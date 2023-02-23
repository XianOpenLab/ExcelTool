import openai

openai.api_key = "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"
response = openai.Image.create(
  prompt="科幻日漫风格的头像",
  n=1,
  size="1024x1024"
)
image_url = response['data'][0]['url']
print(image_url)