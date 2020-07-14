emlファイルを元に扱いやすい様にデータを取得します。

Usage:
    $ python -m eml2txt foo.eml bar.eml ...  # => Dump to foo.txt, bar.txt
    $ python -m eml2txt *.eml -  # => Concat whole eml and dump to STDOUT
