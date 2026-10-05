import io, json, os
S=os.path.dirname(os.path.abspath(__file__))
base=os.path.join(os.path.dirname(S),'')   # תיקיית הדף: זו שמעל qa/
html=io.open(base+'index.html',encoding='utf-8').read()
gs=io.open(base+'apps-script.gs',encoding='utf-8').read()
h=io.open(S+'/qa_harness.js',encoding='utf-8').read()
inj='<script id="gsSrc" type="application/json">'+json.dumps(gs).replace('</','<\\/')+'</script>\n<script>\n'+h.replace('</script','<\\/script')+'\n</script>\n'
assert html.count('</body>')==1
io.open(S+'/qa.html','w',encoding='utf-8').write(html.replace('</body>',inj+'</body>'))
print('built', len(html), '->', os.path.getsize(S+'/qa.html'))
