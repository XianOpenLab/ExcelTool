import openai

openai.api_key = "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"
response = openai.Image.create(
  prompt="科幻感的动漫头像",
  n=1,
  size="1024x1024"
)
image_url = response['data'][0]['url']
print(image_url)