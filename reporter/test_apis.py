import os, json, requests
from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))

print("=== TEST META ===")
try:
    url = f"https://graph.facebook.com/v19.0/{os.getenv('META_AD_ACCOUNT_ID')}/insights"
    params = {
        "level": "campaign",
        "fields": "campaign_name,spend",
        "time_range": json.dumps({"since": "2026-04-01", "until": "2026-04-09"}),
        "access_token": os.getenv("META_ACCESS_TOKEN"),
    }
    r = requests.get(url, params=params)
    print("Status:", r.status_code)
    print("Respuesta:", r.text[:500])
except Exception as e:
    print("Error:", e)

print()
print("=== TEST GOOGLE ADS ===")
try:
    from google.ads.googleads.client import GoogleAdsClient
    client = GoogleAdsClient.load_from_dict({
        "developer_token":   os.getenv("GOOGLE_DEVELOPER_TOKEN"),
        "client_id":         os.getenv("GOOGLE_CLIENT_ID"),
        "client_secret":     os.getenv("GOOGLE_CLIENT_SECRET"),
        "refresh_token":     os.getenv("GOOGLE_REFRESH_TOKEN"),
        "login_customer_id": os.getenv("GOOGLE_MCC_ID").replace("-", ""),
        "use_proto_plus":    True,
    })
    customer_id = os.getenv("GOOGLE_CUSTOMER_ID").replace("-", "")
    ga_service  = client.get_service("GoogleAdsService")
    query = """
        SELECT campaign.name, metrics.cost_micros
        FROM campaign
        WHERE campaign.status = 'ENABLED'
          AND segments.date BETWEEN '2026-04-01' AND '2026-04-09'
        ORDER BY metrics.cost_micros DESC
        LIMIT 5
    """
    stream = ga_service.search_stream(customer_id=customer_id, query=query)
    count = 0
    for batch in stream:
        for result in batch.results:
            cost = result.metrics.cost_micros / 1_000_000
            print(f"  {result.campaign.name} -> ${cost:.2f}")
            count += 1
    print(f"Google Ads OK - {count} campañas")
except Exception as e:
    print("Error:", e)

print()
print("=== TEST GA4 ===")
try:
    from google.analytics.data_v1beta import BetaAnalyticsDataClient
    from google.analytics.data_v1beta.types import RunReportRequest, Dimension, Metric, DateRange
    from google.oauth2 import service_account
    creds = service_account.Credentials.from_service_account_file(
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "ga4-key.json"),
        scopes=["https://www.googleapis.com/auth/analytics.readonly"]
    )
    client = BetaAnalyticsDataClient(credentials=creds)
    req = RunReportRequest(
        property=f"properties/{os.getenv('GA4_PROPERTY_ID')}",
        dimensions=[Dimension(name="sessionSourceMedium")],
        metrics=[Metric(name="sessions")],
        date_ranges=[DateRange(start_date="2026-04-01", end_date="2026-04-09")],
    )
    resp = client.run_report(req)
    print("GA4 OK - filas:", len(resp.rows))
    for row in resp.rows[:3]:
        print(" ", row.dimension_values[0].value, "->", row.metric_values[0].value)
except Exception as e:
    print("Error:", e)
