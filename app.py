from customate_core import *
import json, secrets, threading, webbrowser, socket, ipaddress, html
from urllib.request import Request, build_opener, HTTPRedirectHandler
from urllib.parse import urlparse, urljoin
from html.parser import HTMLParser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
TOKEN=secrets.token_urlsafe(24)
DOC=None
PARTS=[]
MAX_BODY=40*1024*1024

def public_url(url):
 u=urlparse(url)
 if u.scheme not in ('http','https') or not u.hostname or u.username or u.password:raise ValueError('Invalid source URL.')
 for result in socket.getaddrinfo(u.hostname,u.port or (443 if u.scheme=='https' else 80),type=socket.SOCK_STREAM):
  if not ipaddress.ip_address(result[4][0]).is_global:raise ValueError('The source must be a public internet address.')
 return url

class SafeRedirect(HTTPRedirectHandler):
 def redirect_request(self,req,fp,code,msg,headers,newurl):
  public_url(newurl)
  return super().redirect_request(req,fp,code,msg,headers,newurl)

def fetch_public(url,limit=6*1024*1024,timeout=6):
 public_url(url)
 req=Request(url,headers={'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/130.0 Safari/537.36'})
 with build_opener(SafeRedirect).open(req,timeout=timeout) as r:
  content=r.read(limit+1)
  if len(content)>limit:raise ValueError('Source file is too large.')
  return content

class ProductPage(HTMLParser):
 def __init__(self):super().__init__();self.images=[];self.links=[];self.title='';self.in_title=False
 def handle_starttag(self,tag,attrs):
  a=dict(attrs)
  if tag=='title':self.in_title=True
  if tag=='meta' and a.get('property',a.get('name','')).lower() in ('og:image','og:image:url','twitter:image','twitter:image:src') and a.get('content'):self.images.append(a['content'])
  if tag=='img' and a.get('id') in ('landingImage','imgBlkFront','main-image'):
   if a.get('data-old-hires'):self.images.append(a['data-old-hires'])
   if a.get('data-a-dynamic-image'):
    try:
     variants=json.loads(a['data-a-dynamic-image'])
     self.images.extend(sorted(variants,key=lambda u:variants[u][0]*variants[u][1],reverse=True))
    except (ValueError,TypeError,IndexError):pass
   if a.get('src','').startswith(('http://','https://')):self.images.append(a['src'])
  if a.get('data-lpage'):self.links.append(a['data-lpage'])
  if tag=='a' and a.get('href'):self.links.append(a['href'])
 def handle_endtag(self,tag):
  if tag=='title':self.in_title=False
 def handle_data(self,data):
  if self.in_title:self.title+=data

def product_photo(url,pn):
 raw=fetch_public(url);text=raw.decode('utf-8','replace');page=ProductPage();page.feed(text)
 norm=lambda x:re.sub(r'[^A-Z0-9]','',x.upper())
 # The requested PN must occur in the product title or URL, rather than unrelated recommendations.
 key=norm(pn)
 if key not in norm(page.title) and key not in norm(unquote(url)):
  raise ValueError('Product page P/N could not be confirmed.')
 for candidate in page.images[:4]:
  try:
   image_url=urljoin(url,html.unescape(candidate));data=fetch_public(image_url)
   im=Image.open(io.BytesIO(data));im.verify();im=Image.open(io.BytesIO(data)).convert('RGB')
   if im.width<80 or im.height<80:continue
   im.thumbnail((1400,1000));b=io.BytesIO();im.save(b,'PNG')
   return {'image':'data:image/png;base64,'+base64.b64encode(b.getvalue()).decode(),'source':url,'image_url':image_url,'status':'Product image retrieved. P/N found in source title or URL; review image and description.'}
  except Exception:continue
 raise ValueError('The product image could not be downloaded.')

def import_image_url(url,pn):
 raw=fetch_public(url)
 try:
  im=Image.open(io.BytesIO(raw));im.verify();im=Image.open(io.BytesIO(raw)).convert('RGB');im.thumbnail((1400,1000));b=io.BytesIO();im.save(b,'PNG')
  return {'image':'data:image/png;base64,'+base64.b64encode(b.getvalue()).decode(),'source':url,'image_url':url,'status':'Image imported from URL. Please review.'}
 except Exception:return product_photo(url,pn)

class Handler(BaseHTTPRequestHandler):
 def log_message(self,*args):pass
 def send(self,status,data,kind='application/json'):
  self.send_response(status);self.send_header('Content-Type',kind);self.send_header('Content-Length',str(len(data)));self.send_header('Cache-Control','no-store');self.end_headers();self.wfile.write(data)
 def do_GET(self):
  if urlparse(self.path).path!='/':return self.send(404,b'{}')
  self.send(200,HTML.replace('__TOKEN__',TOKEN).encode(),'text/html; charset=utf-8')
 def do_POST(self):
  global DOC,PARTS
  if self.headers.get('X-App-Token')!=TOKEN:return self.send(403,b'{}')
  try:
   n=int(self.headers.get('Content-Length','0'))
   if not 0<n<=MAX_BODY:raise ValueError('Request size exceeds 40 MB.')
   body=self.rfile.read(n)
   if self.path=='/upload':
    DOC,PARTS=parse_doc(body);self.send(200,json.dumps(PARTS).encode())
   elif self.path=='/import-image-url':
    req=json.loads(body);self.send(200,json.dumps(import_image_url(str(req.get('url','')),str(req.get('pn','')))).encode())
   elif self.path=='/product-image':
    req=json.loads(body);self.send(200,json.dumps(product_photo(str(req.get('url','')),str(req.get('pn','')))).encode())
   elif self.path=='/export':
    if DOC is None:raise ValueError('Upload a Word document first.')
    self.send(200,export_doc(DOC,PARTS,json.loads(body)), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
   else:self.send(404,b'{}')
  except Exception as e:self.send(400,json.dumps({'error':str(e)}).encode())

HTML='<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>CUSTOMATE</title><style>\n@import url(\'https://fonts.googleapis.com/css2?family=Titillium+Web:wght@400;600;700&display=swap\');\n[hidden]{display:none!important}*{box-sizing:border-box}body{font:16px \'Titillium Web\',Arial,sans-serif;background:#f2f5f7;color:#142d4d;margin:0}header{background:#102b52;color:white;padding:30px max(24px,calc((100vw - 1050px)/2))}h1{margin:0 0 8px;font-size:30px}main{max-width:1050px;margin:auto;padding:25px}.box{background:white;border-radius:16px;padding:24px;margin:18px 0;border:1px solid #dde6e8}button,.link{background:#173e70;color:white;border:0;border-radius:8px;padding:12px 18px;cursor:pointer;text-decoration:none;display:inline-block}button:disabled{opacity:.45;cursor:default}input,select,textarea{font:inherit}textarea{width:100%;min-height:180px;padding:12px;border:1px solid #c4d2d5;border-radius:8px}label{display:block;margin:15px 0 8px}img{max-width:100%;max-height:260px;object-fit:contain}small,.muted{color:#526b72}.grid{display:grid;grid-template-columns:1fr 2fr;gap:24px}.toolbar{display:flex;gap:12px;align-items:center;flex-wrap:wrap}select,input[type=text]{padding:10px;border:1px solid #c4d2d5;border-radius:8px;max-width:100%}.notice{color:#855515;background:#fff4dc;padding:14px;border-radius:8px}.error{color:#a32121}a{color:#173e70}#status{margin:12px 0}.dropzone{border:2px dashed #8aa1be;border-radius:12px;padding:16px;text-align:center;min-height:150px;outline:none}.dropzone:focus,.dropzone.dragover{border-color:#173e70;background:#eaf1fc}.dropzone img{display:block;margin:auto}.dropzone img[hidden]{display:none}@media(max-width:700px){.grid{grid-template-columns:1fr}}\n</style><header><h1>CUSTOMATE</h1><div>Customs Document Automation</div><div>Upload Word → Review results → Download Word</div></header><main>\n<section class="box"><h2>1. Upload document</h2><p class="muted">Choose a Word document with Part Number, Description, Function, and Detail Picture rows. Letter details are preserved.</p><input type="file" id="doc" accept=".docx"><div id="status" role="status"></div></section>\n<section id="review" hidden><div class="toolbar"><h2>2. Review all items</h2><span id="progress"></span><button onclick="saveSession()">Save progress</button><label style="margin:0">Restore progress <input type="file" id="session" accept=".json"></label></div><p class="notice">Use Search images to find a suitable photo, then paste or drag it into the image box. Review the Indonesian descriptions before downloading. Word font follows your template; English part names and material terms are italicized. The interface is in English; document functions are in Indonesian. Review all drafts before downloading.</p><div class="toolbar"><button id="prev" onclick="move(-1)">← Previous</button><select id="picker" onchange="go(Number(this.value))"></select><button id="next" onclick="move(1)">Next →</button></div><div id="card" class="box"></div><section class="box"><h2>3. Download document</h2><p>Review the results. Missing images will remain blank in the Word document.</p><button id="download" disabled onclick="downloadDoc()">Download Word</button></section></section></main>\n<script>\nlet parts=[],idx=0,fileName=\'Surat_Barang\';const $=id=>document.getElementById(id);const esc=s=>String(s).replace(/[&<>"\']/g,c=>({\'&\':\'&amp;\',\'<\':\'&lt;\',\'>\':\'&gt;\',\'"\':\'&quot;\',"\'":\'&#39;\'}[c]));const mats=[\'metal alloy\',\'engineering plastic\',\'polymer berkekuatan tinggi\',\'rubber dan synthetic polymer\',\'gasket sheet berbasis compressed fiber\',\'carbon\',\'glass\'];const materialLabels={\'polymer berkekuatan tinggi\':\'High-strength polymer\',\'rubber dan synthetic polymer\':\'Rubber and synthetic polymer\',\'gasket sheet berbasis compressed fiber\':\'Compressed fiber gasket sheet\'};\nasync function api(path,body){const r=await fetch(path,{method:\'POST\',headers:{\'X-App-Token\':\'__TOKEN__\'},body});if(!r.ok){const e=await r.json();throw Error(e.error||\'Unable to process the file.\')}return r;}\n$(\'doc\').onchange=async e=>{let f=e.target.files[0];if(!f)return;generation++;$(\'status\').textContent=\'Reading item list…\';try{parts=await(await api(\'/upload\',await f.arrayBuffer())).json();fileName=f.name.replace(/\\.docx$/i,\'\');idx=0;$(\'review\').hidden=false;render();$(\'status\').textContent=`${parts.length} items loaded. Add images using search links, paste, or drag-and-drop.`;}catch(e){$(\'status\').textContent=e.message;}};\nfunction render(){\n$(\'prev\').hidden=true;$(\'next\').hidden=true;$(\'picker\').hidden=true;\n$(\'card\').innerHTML=parts.map((p,i)=>`<section style="border-bottom:1px solid #ddd;padding:20px 0"><h3>${i+1}. ${esc(p.pn)} — ${esc(p.name)}</h3><div class="grid"><div><div class="dropzone" data-drop="${i}" tabindex="0" role="group" aria-label="Image box for ${esc(p.pn)}">${p.image?`<img id="img${i}" src="${esc(p.image)}" alt="Part image">`:`<img id="img${i}" hidden><p id="missing${i}" class="notice">Image not available</p>`}<p class="muted">Drop an image here, or click this box and press Ctrl+V.</p></div><p><a target="_blank" rel="noopener noreferrer" href="https://www.google.com/search?tbm=isch&q=${encodeURIComponent(p.pn+\' \'+p.name)}">Search images ↗</a></p><p id="auto${i}" class="muted">${esc(p.status||\'Manual image selection\')}</p><p id="links${i}">${sourceLinks(p)}</p><input data-image="${i}" type="file" accept="image/png,image/jpeg,image/webp"><label>Source link<input data-source="${i}" type="text" value="${esc(p.source)}" placeholder="Product page URL"></label><button type="button" data-fetch="${i}">Get image from this link</button></div><div><label>Material (draft — check against image)<select data-material="${i}"><option value="">Select material</option>${mats.map(m=>`<option value="${esc(m)}" ${m===p.material?\'selected\':\'\'}>${materialLabels[m]||m}</option>`).join(\'\')}</select></label><label>Function<textarea data-function="${i}">${esc(p.function)}</textarea></label></div></div></section>`).join(\'\');\n$(\'card\').querySelectorAll(\'[data-function]\').forEach(el=>el.oninput=()=>{parts[Number(el.dataset.function)].function=el.value;update();});\n$(\'card\').querySelectorAll(\'[data-fetch]\').forEach(el=>el.onclick=async()=>{const i=Number(el.dataset.fetch),p=parts[i];el.disabled=true;el.textContent=\'Fetching image…\';try{const r=await(await api(\'/product-image\',JSON.stringify({url:p.source,pn:p.pn}))).json();Object.assign(p,r);$(\'img\'+i).src=p.image;$(\'img\'+i).hidden=false;if($(\'missing\'+i))$(\'missing\'+i).hidden=true;$(\'auto\'+i).textContent=p.status;$(\'links\'+i).innerHTML=sourceLinks(p);update();}catch(e){$(\'auto\'+i).textContent=e.message;}finally{el.disabled=false;el.textContent=\'Get image from this link\';}});\n$(\'card\').querySelectorAll(\'[data-source]\').forEach(el=>el.oninput=()=>parts[Number(el.dataset.source)].source=el.value);\n$(\'card\').querySelectorAll(\'[data-material]\').forEach(el=>el.onchange=()=>{const i=Number(el.dataset.material),p=parts[i];const previous=p.material;p.material=el.value;if(p.material&&previous)p.function=p.function.replace(previous,p.material);$(\'card\').querySelector(`[data-function="${i}"]`).value=p.function;update();});\n$(\'card\').querySelectorAll(\'[data-image]\').forEach(el=>el.onchange=()=>{const f=el.files[0];if(f)setItemImage(Number(el.dataset.image),f);});\n$(\'card\').querySelectorAll(\'[data-drop]\').forEach(el=>{\n const i=Number(el.dataset.drop);\n el.onclick=()=>el.focus();\n el.ondragover=e=>{e.preventDefault();el.classList.add(\'dragover\');};\n el.ondragleave=()=>el.classList.remove(\'dragover\');\n el.ondrop=async e=>{e.preventDefault();el.classList.remove(\'dragover\');const f=Array.from(e.dataTransfer.files).find(f=>f.type.startsWith(\'image/\'));if(f){await setItemImage(i,f);return;}const url=e.dataTransfer.getData(\'text/uri-list\')||e.dataTransfer.getData(\'text/plain\');if(/^https?:\\/\\//i.test(url.trim())){await importImageUrl(i,url.trim().split(\'\\n\')[0]);return;}$(\'auto\'+i).textContent=\'Drop an image file, or use Copy image and paste.\';};\n el.onpaste=async e=>{const item=Array.from(e.clipboardData.items).find(x=>x.type.startsWith(\'image/\'));if(item){e.preventDefault();await setItemImage(i,item.getAsFile());return;}const url=e.clipboardData.getData(\'text/plain\').trim();if(/^https?:\\/\\//i.test(url)){e.preventDefault();await importImageUrl(i,url);}};\n});\nupdate();}\n\n\nasync function setItemImage(i,file){if(!file||!file.type.startsWith(\'image/\')){alert(\'Choose an image file.\');return;}if(file.size>10*1024*1024){alert(\'Maximum image size: 10 MB.\');return;}try{const bm=await createImageBitmap(file),cv=document.createElement(\'canvas\'),sc=Math.min(1,1400/bm.width,1000/bm.height);cv.width=Math.round(bm.width*sc);cv.height=Math.round(bm.height*sc);cv.getContext(\'2d\').drawImage(bm,0,0,cv.width,cv.height);parts[i].image=cv.toDataURL(\'image/png\');parts[i].image_url=\'\';$(\'img\'+i).src=parts[i].image;$(\'img\'+i).hidden=false;if($(\'missing\'+i))$(\'missing\'+i).hidden=true;$(\'auto\'+i).textContent=\'Image added successfully.\';if($(\'links\'+i))$(\'links\'+i).innerHTML=sourceLinks(parts[i]);update();bm.close();}catch(e){alert(\'Unable to read image. Use PNG, JPG, or WebP.\');}}\nasync function importImageUrl(i,url){$(\'auto\'+i).textContent=\'Fetching image…\';try{const r=await(await api(\'/import-image-url\',JSON.stringify({url,pn:parts[i].pn}))).json();Object.assign(parts[i],r);$(\'img\'+i).src=r.image;$(\'img\'+i).hidden=false;if($(\'missing\'+i))$(\'missing\'+i).hidden=true;$(\'auto\'+i).textContent=r.status;$(\'links\'+i).innerHTML=sourceLinks(parts[i]);const field=$(\'card\').querySelector(`[data-source="${i}"]`);if(field)field.value=parts[i].source||\'\';update();}catch(e){$(\'auto\'+i).textContent=\'Unable to import link. Use Copy image → Ctrl+V, or drop a saved image file.\';}}\n\nfunction update(){const missing=parts.filter(p=>!p.image).length,empty=parts.filter(p=>!p.function.trim()).length;$(\'progress\').textContent=`${parts.length} item · ${missing} missing images · ${empty} empty functions`;$(\'download\').disabled=empty>0;}\n\nfunction sourceLinks(p){let result=\'\';for(const [label,url] of [[\'Image source\',p.source],[\'Open image\',p.image_url]]){if(url&&/^https?:\\/\\//i.test(url))result+=`<a target="_blank" rel="noopener noreferrer" href="${esc(url)}">${label} ↗</a> &nbsp; `;}return result;}\nlet generation=0;\nfunction move(n){go(idx+n)}function go(n){idx=n;render();}function saveBlob(b,name){const u=URL.createObjectURL(b);const a=document.createElement(\'a\');a.href=u;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(u),2000);}\nfunction saveSession(){saveBlob(new Blob([JSON.stringify({fileName,parts})],{type:\'application/json\'}),fileName+\'_pekerjaan.json\');}\n$(\'session\').onchange=async e=>{try{const saved=JSON.parse(await e.target.files[0].text());if(!Array.isArray(saved.parts)||saved.parts.length!==parts.length||saved.parts.some((p,i)=>p.pn!==parts[i].pn||p.name!==parts[i].name))throw Error(\'Saved progress does not match the uploaded Word document.\');for(let i=0;i<parts.length;i++){const s=saved.parts[i];if(s.image&&!/^data:image\\/(png|jpeg|webp);base64,/.test(s.image))throw Error(\'Invalid image format.\');Object.assign(parts[i],{function:String(s.function||\'\'),material:mats.includes(s.material)?s.material:\'\',image:s.image||\'\',source:String(s.source||\'\'),reviewed:false});}idx=0;render();}catch(e){alert(e.message);}};\nasync function downloadDoc(){const missing=parts.filter(p=>!p.image).length;if(missing&&!confirm(`${missing} items have no images. Download Word with blank image sections?`))return;const b=$(\'download\');b.disabled=true;b.textContent=\'Building Word document…\';try{const r=await api(\'/export\',JSON.stringify(parts));saveBlob(await r.blob(),fileName+\'_Terisi.docx\');}catch(e){alert(e.message);}finally{b.textContent=\'Download Word\';update();}}\n</script></html>'

if __name__=='__main__':
    server=ThreadingHTTPServer(('127.0.0.1',0),Handler)
    url=f'http://127.0.0.1:{server.server_address[1]}'
    print('CUSTOMATE: '+url+'\nKeep this console open while using the application.')
    threading.Timer(1,lambda:webbrowser.open(url)).start()
    try:server.serve_forever()
    except KeyboardInterrupt:pass
    finally:server.server_close()
