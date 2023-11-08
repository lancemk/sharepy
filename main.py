import sharepy
import os
import re
domain_base = "entuedu-my.sharepoint.com"
sharpoint_root = domain_base + "/personal/bochuan-mok_e_ntu_edu_sg/_layouts/15/onedrive.aspx?id="

acc_root = "%2Fpersonal%2Fbochuan-mok_e_ntu_edu_sg%2F"
context_root = "business%2FinPlots%2F"
cf_path = "__corporate%20function__%2F"

query = "%2Fpersonal%2Fbochuan-mok_e_ntu_edu_sg%2FDocuments%2Fbusiness%2FinPlots%2F__corporate%20function__%2F"


def fmt_urlfy(chunk):
    res = ""
    for c in chunk:
        if c == '_':
            res = res + '%5F'
        elif c == '-':
            res = res + '%2D'         
        elif c == '/':
            res = res + '%2F'        
        elif c == ' ':
            res = res + '%20'
        else:
            res = res + c
    return res


if os.path.isfile(os.path.join(".", "sp-session.pkl")):
    s = sharepy.load()
else:
    s = sharepy.connect(domain_base)
    s.save()


def get_lists():
    """ "application/xml" """
    url = "https://" + domain_base + "/personal/bochuan-mok_e_ntu_edu_sg/_api/web/lists"
    print(url)
    r = s.get(url,
        headers={"Accept": "application/atom+xml"})
    
    print(r.content)


def get_file():
    file = "human resource policy.xlsx"
    # target_url = "https://" + sharpoint_root + acc_root + context_root + cf_path + file.replace(" ", "%20")
    target_url = "https://" + sharpoint_root + fmt_urlfy(query) + file.replace(" ", "%20")
    print("{} -> {}".format(target_url, target_url))
    r = s.getfile(target_url, filename=f'download/{file}')
    print(r)


# "https://entuedu-my.sharepoint.com/personal/bochuan-mok_e_ntu_edu_sg/_layouts/15/onedrive.aspx?id=%2Fpersonal%2Fbochuan%2Dmok%5Fe%5Fntu%5Fedu%5Fsg%2FDocuments%2Fbusiness%2FinPlots%2F%5F%5Fcorporate%20function%5F%5F&view=0"
# "https://entuedu-my.sharepoint.com/personal/bochuan-mok_e_ntu_edu_sg/_layouts/15/onedrive.aspx?id=%2Fpersonal%2Fbochuan%2Dmok%5Fe%5Fntu%5Fedu%5Fsg%2FDocuments%2Fbusiness%2FinPlots%2F%5F%5Fcorporate%20function%5F%5F%2Fhuman%20resource%20policy.xlsx"


# https://entuedu-my.sharepoint.com/personal/bochuan-mok_e_ntu_edu_sg/_api/web/lists
# https://entuedu-my.sharepoint.com/personal/bochuan-mok_e_ntu_edu_sg/api/web/lists
get_lists()