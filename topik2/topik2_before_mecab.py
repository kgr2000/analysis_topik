import codecs
from konlpy.tag import Mecab
mecab = Mecab()

txt = codecs.open('data_test.txt', 'r', encoding="utf-8")

for line in txt :
    print(mecab.pos(line))