#!/usr/bin/env python3
# coding:utf-8
"""
emlファイルを元に扱いやすい様にデータを取得します。

Usage:
    python -m eml2txt foo.eml bar.eml ...
"""
import sys
import email
from email.header import decode_header

class MailParser(object):
    """
    メールファイルのパスを受け取り、それを解析するクラス
    """

    def __init__(self, mail_file_path):
        self.mail_file_path = mail_file_path
        # emlファイルからemail.message.Messageインスタンスの取得
        with open(mail_file_path, 'rb') as email_file:
            self.email_message = email.message_from_bytes(email_file.read())
        self.subject = None
        self.to_address = None
        self.cc_address = None
        self.from_address = None
        self.body = ""
        # 添付ファイル関連の情報
        # {name: file_name, data: data}
        self.attach_file_list = []
        # emlの解釈
        self._parse()

    def get_attr_data(self):
        """
        メールデータの取得
        """
        result = """\
FROM: {}
TO: {}
CC: {}
-----------------------
SUBJECT: {}
BODY:
{}
-----------------------
ATTACH_FILE_NAME:
{}
""".format(
            self.from_address,
            self.to_address,
            self.cc_address,
            self.subject,
            self.body,
            ",".join([ x["name"] for x in self.attach_file_list])
        )
        return result


    def _parse(self):
        """
        メールファイルの解析
        __init__内で呼び出している
        """
        self.subject = self._get_decoded_header("Subject")
        self.to_address = self._get_decoded_header("To")
        self.cc_address = self._get_decoded_header("Cc")
        self.from_address = self._get_decoded_header("From")

        # メッセージ本文部分の処理
        for part in self.email_message.walk():
            # ContentTypeがmultipartの場合は実際のコンテンツはさらに
            # 中のpartにあるので読み飛ばす
            if part.get_content_maintype() == 'multipart':
                continue
            # ファイル名の取得
            attach_fname = part.get_filename()
            # ファイル名がない場合は本文のはず
            if not attach_fname:
                charset = str(part.get_content_charset())
                if charset:
                    self.body += part.get_payload(decode=True).decode(charset, errors="replace")
                else:
                    self.body += part.get_payload(decode=True)
            else:
                # ファイル名があるならそれは添付ファイルなので
                # データを取得する
                self.attach_file_list.append({
                    "name": attach_fname,
                    "data": part.get_payload(decode=True)
                })

    def _get_decoded_header(self, key_name):
        """
        ヘッダーオブジェクトからデコード済の結果を取得する
        """
        ret = ""

        # 該当項目がないkeyは空文字を戻す
        raw_obj = self.email_message.get(key_name)
        if raw_obj is None:
            return ""
        # デコードした結果をunicodeにする
        for fragment, encoding in decode_header(raw_obj):
            if not hasattr(fragment, "decode"):
                ret += fragment
                continue
            # encodeがなければとりあえずUTF-8でデコードする
            if encoding:
                try:
                    ret += fragment.decode(encoding)
                except UnicodeDecodeError:
                    ret += fragment.decode(encoding, errors='ignore')
            else:
                ret += fragment.decode("UTF-8")
        return ret

if __name__ == "__main__":
    if len(sys.argv) <2:  # No args
        print(__doc__)
        sys.exit(1)
    elif sys.argv[1] == '-h' or sys.argv[1] == '--help':  # Show help
        print(__doc__)
        sys.exit(0)
    elif '-' in sys.argv:  # Dump messages to STDOUT
        sys.argv.remove('-')
        for filename in sys.argv[1:]:
            result = MailParser(filename).get_attr_data()
            print(result)
    else:
        for filename in sys.argv[1:]:  # Dump messages to TEXT
            parser = MailParser(filename)
            subject = parser.subject.replace(':', '').replace(' ', '')
            filename = f'{subject}.txt'
            result = parser.get_attr_data()
            with open(filename, 'w') as f:
                f.write(result)
