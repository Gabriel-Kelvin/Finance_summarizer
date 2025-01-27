from openai import AzureOpenAI
import collections

collections.Callable = collections.abc.Callable

client = AzureOpenAI(

    api_key="3txjyxDMw3BbClF4YSqpHwGPsOfoSlRFl22zkrvU2KiIv6tcFdwcJQQJ99ALAC77bzfXJ3w3AAABACOG41Xj",

    api_version="2024-08-01-preview",

    azure_endpoint="https://neohackathon01-2024.openai.azure.com/"

)

deployment_name = 'gpt-4o'

test_prompt = "Write a tagline for neostats hackathon"


def test_model():
    try:

        response = client.chat.completions.create(

            model=deployment_name,

            messages=[{"role": "user", "content": test_prompt}],

            temperature=0.7,

        )

        print("Model response:", response.choices[0].message.content)

    except Exception as e:

        print("Error:", str(e))


test_model()