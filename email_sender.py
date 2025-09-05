#!/usr/bin/env python
# coding: utf-8

# In[3]:


import win32com.client as win32
from pathlib import Path

def send_outlook_email(
    subject,
    to_recipients,
    cc_recipients=None,
    html_body="",
    signature="",
    attachment_folder=None,
    display_before_send=False,
    font_family="Microsoft JhengHei"  
):
    """
    使用 Outlook 發送電子郵件（支援 HTML 內文、HTML 簽名、資料夾附件，可自訂字型）

    參數：
    - subject: 郵件主旨
    - to_recipients: 收件者清單 (list)
    - cc_recipients: 副本清單 (list，可省略)
    - html_body: 郵件 HTML 內文
    - signature: 郵件結尾簽名 (HTML 格式，支援 <br>)
    - attachment_folder: 要附加的資料夾 (會加入底下所有檔案)
    - display_before_send: True → 先顯示再手動送出；False → 直接寄出
    - font_family: 指定字型名稱（預設 Microsoft JhengHei）
    """

    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # 0 = olMailItem

    # === 設定基本資料 ===
    mail.Subject = subject
    mail.To = "; ".join(to_recipients)
    if cc_recipients:
        mail.CC = "; ".join(cc_recipients)

    # === HTML 信件內文 ===
    style = f"""
    <style>
        body {{
            font-family: '{font_family}', sans-serif;
            font-size: 12pt;
        }}
    </style>
    """

    # === 組合內文+簽名 ===
    if signature:
        final_html = f"{style}<body>{html_body}<br><br>{signature}</body>"
    else:
        final_html = f"{style}<body>{html_body}</body>"
    mail.HTMLBody = final_html

    # === 加入附件 ===
    if attachment_folder:
        folder_path = Path(attachment_folder)
        if folder_path.is_dir():
            for file in folder_path.iterdir():
                if file.is_file():
                    mail.Attachments.Add(str(file))

    # === 顯示或直接送出 ===
    if display_before_send:
        mail.Display(True)
    else:
        mail.Send()


# In[5]:





# In[ ]:




