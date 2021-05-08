import imaplib
import email.utils
import time
import base64

# 参考：https://stackoverflow.com/questions/12776679/imap-folder-path-encoding-imap-utf-7-for-python
def utf7_encode(s):
    """"Encode a string into RFC2060 aka IMAP UTF7"""
    s = s.replace('&', '&-')
    unipart = out = ''
    for c in s:
        if 0x20 <= ord(c) <= 0x7f:
            if unipart != '':
                out += '&' + base64.b64encode(unipart.encode('utf-16-be')).decode('ascii').rstrip('=') + '-'
                unipart = ''
            out += c
        else:
            unipart += c
    if unipart != '':
        out += '&' + base64.b64encode(unipart.encode('utf-16-be')).decode('ascii').rstrip('=') + '-'
    return out

# 接続情報
imap_host = 'outlook.office365.com'
username = 'username@example.com'
password = 'xxxxxxxx'
mailbox = '送信済みアイテム'

# メールサーバーに接続
imap = imaplib.IMAP4_SSL(imap_host)
imap.login(username, password)

# 受信ボックスを選択
imap.select(utf7_encode(mailbox))

# 今日送信したメールを取得
typ, list = imap.search(None, '(ON {0})'.format(time.strftime('%d-%b-%Y')))

# 送信時刻の一覧を取得
sent_times = []
for num in list[0].split():
    typ, data = imap.fetch(num, '(RFC822)')
    raw_email = data[0][1]
    message = email.message_from_string(raw_email.decode('utf-8'))
    date = email.utils.parsedate(message.get('Date'))
    sent_times.append('{:02d}時{:02d}分{:02d}秒'.format(date[3], date[4], date[5]))

# 最初と最後の送信時刻を出力
print(sent_times[0])
print(sent_times[len(sent_times) - 1])

# メールサーバーへの接続を切断
imap.close()
imap.logout()
