key = "af3378c267ba41c9a94eedc57146de03"
endpoint = "https://251094676.cognitiveservices.azure.com/"
#change this key and endpoint if you have your own azure congnitive service

from azure.ai.textanalytics import TextAnalyticsClient
from azure.core.credentials import AzureKeyCredential

# Authenticate the client using your key and endpoint 
def authenticate_client():
    ta_credential = AzureKeyCredential(key)
    text_analytics_client = TextAnalyticsClient(
            endpoint=endpoint, 
            credential=ta_credential)
    return text_analytics_client

client = authenticate_client()