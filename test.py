from sharepoint import SharePointSite
from sharepoint.auth import PreemptiveBasicAuthHandler
from ntlm import HTTPNtlmAuthHandler
from urllib import request

# "https://intermountainhealth.sharepoint.com/sites/DRMCSleepCenter/Shared Documents/IHC-SLEEP STUDIES-Y-T-D.xlsx"
sharepoint_url = "https://intermountainhealth.sharepoint.com/"
# copy'n'paste from auth.py :)
password_manager = request.HTTPPasswordMgrWithDefaultRealm()
password_manager.add_password(None, sharepoint_url, "pwatkin1", "6bkrVLmuyn$w")
auth_handler = HTTPNtlmAuthHandler.HTTPNtlmAuthHandler(password_manager)

# ph = request.ProxyHandler( { } )
opener = request.build_opener(auth_handler)

site = SharePointSite(sharepoint_url, opener)