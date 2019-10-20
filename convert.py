from bs4 import BeautifulSoup
import os
import getpass
import pypandoc
import requests
import sys
import urllib.parse

# Global Variables
base_url = "https://" + input("Instance (e.g. dev98765): ") + ".service-now.com"
auth_user = input("Username: ")
print("Please note that your password is being recorded even though no visual feedback is shown.")
auth_pass = getpass.getpass("Password: ")
knowledge_base = input("Knowledge Base sys_id: ")

# File Structure
in_path  = './Input (Word)/'
int_path = './Intermediary (Markdown)/'
out_path = './Output (HTML)/'
if not os.path.exists(in_path):
    os.makedirs(in_path)
if not os.path.exists(int_path):
    os.makedirs(int_path)
if not os.path.exists(out_path):
    os.makedirs(out_path)

for os_file in os.listdir(in_path):
    # File Structure
    title = '.'.join(os_file.split(".")[:-1])
    in_file  = in_path + title + '.docx'
    int_file = int_path + title + '.md'
    out_file = out_path + title + '.html'
    # ServiceNow Parameters
    headers = {"Content-Type":"application/json","Accept":"application/json"}
    # Actual Script
    ## Pandoc Section
    output_markdown = pypandoc.convert_file(in_file, 'markdown_strict', extra_args = ['--wrap=none', '--reference-links', '--extract-media=' + int_path])
    output_html = pypandoc.convert_text(output_markdown, to = 'html', format = 'markdown', extra_args = ['--wrap=none', '--reference-links'])
    ## Replace Image src
    imgs = []
    soup = BeautifulSoup(output_html, "lxml")
    for img in soup.find_all("img"):
        src = img['src'].split("/")
        src = int_path + "media/" + src[len(src) - 1]
        imgs.append({"local_source": src, "new_src": '/sys_attachment.do?sys_id=xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'})
        img['src'] = '/sys_attachment.do?sys_id=xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'
    output_html = str(soup)
    ## Import Article into ServiceNow
    sn_data = {
        "kb_knowledge_base": knowledge_base,
        "short_description": title,
        "text": output_html,
    }
    knowledge_response = requests.post(base_url + '/api/now/table/kb_knowledge', auth = (auth_user, auth_pass, ), headers = headers, json = sn_data)
    if knowledge_response.status_code != 201:
        print('Status:', knowledge_response.status_code, 'Headers:', knowledge_response.headers, 'Error Response:', knowledge_response.json())
        exit()
    knowledge_sys_id = knowledge_response.json()['result']['sys_id']
    print("Inserted kb_knowledge.do?sys_id=%s" % (knowledge_sys_id, ))
    ## Create Attachments
    imgs_new = imgs
    n = 0
    for img in imgs:
        img_header = {"Accept":"*/*"}
        mimetype = img['local_source'].split(".")
        mimetype = mimetype[len(mimetype) - 1]
        payload = {'table_name': 'kb_knowledge', 'table_sys_id': knowledge_sys_id}
        files = {
            'file': (img['local_source'], open(img['local_source'], 'rb'),
            'image/' + mimetype,
            {'Expires': '0'})
        }
        img_response = requests.post(url = base_url + '/api/now/attachment/upload', auth = (auth_user, auth_pass, ), headers = img_header, files = files, data = payload)
        if img_response.status_code != 201:
            print('Status:', img_response.status_code, 'Headers:', img_response.headers, 'Error Response:', img_response.json())
            exit()
        attachment_sys_id = img_response.json()['result']['sys_id']
        imgs_new[n]["sys_id"] = "/sys_attachment.do?sys_id=" + attachment_sys_id
        print("Inserted kb_knowledge.do?sys_id=%s's sys_attachment.do?sys_id=%s" % (knowledge_sys_id, attachment_sys_id, ))
        n += 1
    ## Re-Replace Image src
    imgs = []
    soup = BeautifulSoup(output_html, "lxml")
    n = 0
    for img in soup.find_all("img"):
        img['src'] = imgs_new[n]["sys_id"]
        n += 1
    output_html = str(soup)
    ## Re-Upload Knowledge to ServiceNow
    sn_data = {"text": output_html,}
    knowledge_response = requests.put(base_url + '/api/now/table/kb_knowledge/' + knowledge_sys_id, auth = (auth_user, auth_pass, ), headers = headers, json = sn_data)
    if knowledge_response.status_code != 200:
        print('Status:', knowledge_response.status_code, 'Headers:', knowledge_response.headers, 'Error Response:', knowledge_response.json())
        exit()
    knowledge_sys_id = knowledge_response.json()['result']['sys_id']
    print("Updated  kb_knowledge.do?sys_id=%s" % (knowledge_sys_id, ))
