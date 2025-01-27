from azure.ai.formrecognizer import DocumentAnalysisClient

from azure.core.credentials import AzureKeyCredential


def get_content(file_path):
    document_analysis_client = DocumentAnalysisClient(

        endpoint="https://neohack.cognitiveservices.azure.com/",

        credential=AzureKeyCredential(
            "D4Vd1Xljol6cS0q8PoZiQFXz8uH1yYeXCIyhTHU5vdBxBhbeX9JLJQQJ99ALACYeBjFXJ3w3AAALACOGIiVt")
    )

    # Open the file in binary mode

    with open(file_path, "rb") as file:
        poller = document_analysis_client.begin_analyze_document("prebuilt-layout", document=file)

        result = poller.result()

        return result.content


# File path to the PDF document

file_path = "Report.pdf"

# Get and print the content

try:

    content = get_content(file_path)

    print(content)

except Exception as e:

    print(f"Error: {e}")