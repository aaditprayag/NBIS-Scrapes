import re,html,sys
s=open(sys.argv[1],errors='ignore').read()
t=re.sub(r'<script.*?</script>|<style.*?</style>|<svg.*?</svg>','',s,flags=re.S)
t=re.sub(r'<[^>]+>','|',t); t=html.unescape(t); t=re.sub(r'\s*\|[\s|]*','|',t)
print(t)
