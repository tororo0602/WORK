class SplashArtwork:
    def text(self, x, y, text, size=10, color='#A3C1D1', family='Consolas', bold=False, **kw):
            return self.canvas.create_text(x,y,text=text,anchor='nw',fill=color,
                                           font=(family,size,'bold' if bold else 'normal'), **kw)

    def show(self, stage, progress=68, cancel=True):
            if cancel and self.job:
                self.root.after_cancel(self.job)
                self.job=None
            c=self.canvas
            c.delete('all')
            c.create_image(0,0,image=self.image,anchor='nw')
            self.text(32,32,'NC PROGRAM CHECK',11,'#99D5EA',bold=True)
            self.text(32,135,'NCプログラム',22,'#EFF9FF',family='Yu Gothic UI',bold=True)
            self.text(32,174,'確認ツール',22,'#EFF9FF',family='Yu Gothic UI',bold=True)
            if stage == 1:
                c.create_line(34,239,64,239,fill='#63D8FF',width=2)
                self.text(77,228,'起動準備中',11,'#B4C9D4',family='Yu Gothic UI')
                self.text(32,430,'STARTING',9,'#6E99AE')
            else:
                self.text(34,226,'SYSTEM INITIALIZATION',9,'#83CDE5')
                self.text(274,224,f'{progress:03d}%',12,'#BCEFFF',bold=True)
                c.create_rectangle(34,253,322,255,fill='#17313F',outline='')
                c.create_rectangle(34,253,34+288*progress/100,255,fill='#71DEFF',outline='')
                steps=[('CORE.RUNTIME','実行環境',15),('NC.PARSER','NC解析エンジン',35),
                       ('RULE.INDEX','入力規則',58),('UI.PIPELINE','画面の構築',85),('WORKSPACE','起動準備',100)]
                active=next((i for i,s in enumerate(steps) if progress<s[2]),len(steps)-1)
                for i,(name,label,threshold) in enumerate(steps):
                    done=progress>=threshold
                    current=i==active and progress<100
                    color='#BADAE7' if done else '#77DEFF' if current else '#4C697A'
                    self.text(34,273+i*23, name,9,color)
                    self.text(255,273+i*23,'READY' if done else 'LOAD' if current else 'WAIT',9,color)
                line='起動準備が完了しました' if progress==100 else steps[active][1]+'を読み込み中'
                self.text(34,405,line,10,'#AAC5D4',family='Yu Gothic UI')
                self.text(34,433,'INITIALIZATION COMPLETE' if progress==100 else 'BOOT / MODULE SEQUENCE',8,'#5C8CA3')
            self.root.update_idletasks()
