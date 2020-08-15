#!/usr/bin/env python3
# coding:utf-8
"""
emlファイルを元に扱いやすい様にデータを取得します。

Usage:
    $ python -m eml2txt foo.eml bar.eml ...  # => Dump to foo.txt, bar.txt
    $ python -m eml2txt *.eml -  # => Concat whole eml and dump to STDOUT
"""
import sys
import re
from dateutil.parser import parse
import email
from email.header import decode_header

VERSION = "eml2ext v1.0.1r"


class MailParser:
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
        self.date = None
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
DATE: {}
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
""".format(self.date, self.from_address, self.to_address, self.cc_address,
           self.subject, self.body,
           ",".join([x["name"] for x in self.attach_file_list]))
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
        self.date = parse(self._get_decoded_header("Date")).isoformat()

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
                charset = part.get_content_charset()
                if charset:
                    self.body += part.get_payload(decode=True).decode(
                        str(charset), errors="replace")
            else:
                # ファイル名があるならそれは添付ファイルなので
                # データを取得する
                self.attach_file_list.append({
                    "name":
                    attach_fname,
                    "data":
                    part.get_payload(decode=True)
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
            ret += fragment.decode(encoding if encoding else 'UTF-8',
                                   errors='replace')
        return ret

    @staticmethod
    def help(exitcode):
        """Show help"""
        print(__doc__)
        sys.exit(exitcode)

    @staticmethod
    def version():
        """Show version"""
        print(VERSION)
        sys.exit(0)

    @classmethod
    def dump2stdout(cls, argv):
        """Dump messages to STDOUT"""
        argv.remove('-')
        for filename in argv[1:]:
            result = cls(filename).get_attr_data()
            print(result)

    @classmethod
    def dump2txt(cls, argv):
        """Dump messages to TEXT"""
        for filename in argv[1:]:
            parser = cls(filename)
            invalid_str = r"[\\/:*?\"<>|]"
            # Remove invalid text
            subject = re.sub(invalid_str, "", parser.subject)
            result = parser.get_attr_data()
            with open(
                    f'{subject}.txt',
                    'a',  # Append same subject exitst
                    encoding='utf-8') as _f:
                _f.write(result)


def main():
    """Entry point"""
    if len(sys.argv) < 2:  # No args
        MailParser.help(1)
    elif sys.argv[1] == '-v' or sys.argv[1] == '--version':
        MailParser.version()
    elif sys.argv[1] == '-h' or sys.argv[1] == '--help':
        MailParser.help(0)
    elif '-' in sys.argv:
        MailParser.dump2stdout(sys.argv)
    else:
        MailParser.dump2txt(sys.argv)


if __name__ == "__main__":
    main()
