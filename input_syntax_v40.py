"""NC手編集時の書式チェック。機械固有コードの可否・値の意味は判定しない。"""
import re
from functools import lru_cache


class InputSyntaxChecker:
    FUNCTIONS = {'ABS', 'ACOS', 'ASIN', 'ATAN', 'COS', 'SIN', 'TAN', 'SQRT',
                 'EXP', 'LN', 'FIX', 'FUP', 'ROUND', 'BIN', 'BCD', 'ADP'}
    WORDS = FUNCTIONS | {'IF', 'THEN', 'GOTO', 'WHILE', 'DO', 'END',
                         'EQ', 'NE', 'GT', 'GE', 'LT', 'LE', 'AND', 'OR', 'XOR', 'MOD'}
    OPS = {'+', '-', '*', '/', '**', 'EQ', 'NE', 'GT', 'GE', 'LT', 'LE', 'AND', 'OR', 'XOR', 'MOD'}

    SORTED_WORDS = tuple(sorted(WORDS, key=len, reverse=True))

    def __init__(self, line):
        self.line = line
        self.issues = []
        self.tokens = []
        self.i = 0

    def issue(self, start, end, message):
        item = (start, max(start + 1, end), message)
        if item not in self.issues:
            self.issues.append(item)

    def scan(self):
        s = self.line
        i = 0
        while i < len(s):
            c = s[i]
            if c == '(':
                start = i
                i = s.find(')', i + 1)
                if i < 0:
                    self.issue(start, len(s), 'コメントの閉じ括弧がありません')
                    break
                i += 1
                continue
            if c == ')':
                self.issue(i, i + 1, '対応するコメントの開き括弧がありません')
                i += 1
                continue
            if c in ' \t\r\n':
                i += 1
                continue
            if ord(c) > 127:
                self.issue(i, i + 1, '全角文字・半角以外の文字や記号が混入しています')
                i += 1
                continue
            if c.isdigit() or c == '.':
                m = re.match(r'[0-9.]+', s[i:])
                value = m.group()
                self.tokens.append(('number', value, i, i + len(value)))
                if value == '.':
                    self.issue(i, i + 1, '小数点だけで数値がありません')
                elif value.count('.') > 1:
                    self.issue(i, i + len(value), '小数点が複数あります')
                i += len(value)
                continue
            if c.isalpha():
                word = next((w for w in self.SORTED_WORDS
                             if s[i:i + len(w)].upper() == w), None)
                value = word or c.upper()
                self.tokens.append(('word' if word else 'address', value, i, i + len(value)))
                i += len(value)
                continue
            if c in '#[]=+-*/%,:;':
                value = '**' if s[i:i+2] == '**' else c
                self.tokens.append(('symbol', value, i, i + len(value)))
                i += len(value)
                continue
            self.issue(i, i + 1, '解釈できない記号が混入しています')
            i += 1
        self.tokens.append(('eof', '', len(s), len(s)))

    def token(self):
        return self.tokens[self.i]

    def take(self):
        t = self.token()
        if t[0] != 'eof':
            self.i += 1
        return t

    def atom(self, depth=0):
        if depth > 80:
            t = self.take()
            self.issue(t[2], t[3], '式が深すぎるため書式を確認してください')
            return True
        t = self.token()
        if t[1] in ('+', '-'):
            self.take()
            if self.token()[1] in ('+', '-'):
                last = self.take()
                self.issue(t[2], last[3], '符号が重複しています')
            if not self.atom(depth + 1):
                self.issue(t[2], t[3], '符号の後ろに値がありません')
            return True
        if t[0] == 'number':
            self.take()
            return True
        if t[1] == '#':
            self.take()
            n = self.token()
            if n[0] == 'number' and n[1].isdigit():
                self.take()
            elif n[1] in ('[', '#'):
                self.atom(depth + 1)
            else:
                self.issue(t[2], t[3], '変数番号・参照式がありません')
            return True
        if t[1] in self.FUNCTIONS:
            self.take()
            if self.token()[1] == '[':
                self.atom(depth + 1)
                if t[1] == 'ATAN' and self.token()[1] == '/':
                    self.take()
                    if not self.atom(depth + 1):
                        self.issue(t[2], self.token()[2], '関数の引数がありません')
            else:
                self.issue(t[2], t[3], '関数の引数を囲む角括弧がありません')
            return True
        if t[1] == '[':
            self.take()
            if not self.atom(depth + 1):
                self.issue(t[2], t[3], '式の値がありません')
            while self.token()[1] in self.OPS:
                op = self.take()
                if not self.atom(depth + 1):
                    self.issue(op[2], op[3], '演算子の後ろに値がありません')
                    break
            if self.token()[1] == ']':
                self.take()
            else:
                self.issue(t[2], t[3], '式の閉じ括弧がないか、式の書式が不正です')
            return True
        return False

    def run(self):
        self.scan()
        while self.token()[0] != 'eof':
            t = self.token()
            if t[0] == 'address':
                self.take()
                if not self.atom():
                    self.issue(t[2], t[3], f'{t[1]}の指令値がありません')
            elif t[1] == '#':
                self.atom()
                if self.token()[1] == '=':
                    eq = self.take()
                    if not self.atom():
                        self.issue(eq[2], eq[3], '代入する値・式がありません')
                    while self.token()[1] in self.OPS:
                        op = self.take()
                        if not self.atom():
                            self.issue(op[2], op[3], '演算子の後ろに値がありません')
                            break
            elif t[1] in ('IF', 'WHILE'):
                self.take()
                if self.token()[1] == '[':
                    self.atom()
                else:
                    self.issue(t[2], t[3], '条件式がありません')
            elif t[1] in ('GOTO', 'DO', 'END'):
                self.take()
                if not self.atom():
                    self.issue(t[2], t[3], f'{t[1]}の番号がありません')
            elif t[1] == 'THEN':
                self.take()
            elif t[1] == '/' and not self.line[:t[2]].strip():
                self.take()  # ブロックスキップ / または /1～/9
                if self.token()[0] == 'number':
                    self.take()
            elif t[1] in ('%', ':', ';'):
                self.take()
                if t[1] == ':' and self.token()[0] == 'number':
                    self.take()  # プログラム番号の別表記
            else:
                self.take()
                message = ('どの指令にも属さない数値があります' if t[0] == 'number'
                           else '対応する開き角括弧がありません' if t[1] == ']'
                           else '不要な符号・演算子、または入力途中の式があります')
                self.issue(t[2], t[3], message)
        return sorted(self.issues)


@lru_cache(maxsize=32768)
def _cached_syntax(line):
    return tuple(InputSyntaxChecker(line).run())

def check_input_syntax(line):
    # Callers get independent lists; cached values cannot be mutated.
    return list(_cached_syntax(line))
