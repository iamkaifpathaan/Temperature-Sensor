import html
from qa import FAQ, HARD
e=html.escape
f=[]
for i,(tag,q,ask,exp,say,dont) in enumerate(FAQ,1):
    f.append(f'''<div class="faq"><div class="q"><span class="n">{i}.</span>“{e(q)}”<span class="tag">{e(tag)}</span></div><div class="rows">
<div class="h">Really asking</div><div class="v">{e(ask)}</div>
<div class="h">For us</div><div class="v">{e(exp)}</div>
<div class="h say">Say</div><div class="v say quote">{e(say)}</div>
<div class="h dont">Don't promise</div><div class="v">{e(dont)}</div></div></div>''')
h=[]
for i,(q,say,note) in enumerate(HARD,1):
    h.append(f'''<div class="faq"><div class="q"><span class="n">{i}.</span>{e(q)}</div><div class="rows">
<div class="h say">Say</div><div class="v say quote">{e(say)}</div>
<div class="h">Note for us</div><div class="v">{e(note)}</div></div></div>''')
s=open('playbook.html').read().replace('{{FAQ}}','\n'.join(f)).replace('{{HARD}}','\n'.join(h))
open('playbook_out.html','w').write(s)
print(len(FAQ),len(HARD))
