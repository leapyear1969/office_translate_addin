// Local visual QA only: mocks Office, account and translation. Never bundled or deployed.
const express = require('express');
const fs = require('fs');
const path = require('path');
const app = express();
const mock = `<script>
window.OfficeRuntime={auth:{getAccessToken:async()=> 'preview'}};
window.Office={HostType:{Word:'Word'},PlatformType:{PC:'PC',Mac:'Mac',OfficeOnline:'OfficeOnline'},AsyncResultStatus:{Succeeded:'succeeded'},VisibilityMode:{taskpane:'taskpane'},actions:{associate(){}},addin:{showAsTaskpane:async()=>{},onVisibilityModeChanged:async()=>{}},context:{platform:'PC',document:{settings:{get:()=>({target:'zh-Hans'}),set(){},saveAsync(cb){cb({status:'succeeded'})}}}},onReady(cb){document.addEventListener('DOMContentLoaded',()=>cb({host:'Word'}))}};
const range={text:'你好，杰森',load(){},insertText(){}};range.paragraphs={getFirst:()=>({getRange:()=>range})};
const ctx={document:{getSelection:()=>range},sync:async()=>{},trackedObjects:{add(){},remove(){}}};window.Word={run:async(...args)=>args.at(-1)(ctx),InsertLocation:{replace:'replace'}};
window.addEventListener('error',e=>fetch('/preview-error',{method:'POST',body:e.message}));
</script>`;
app.get('/word/taskpane.html',(req,res)=>res.send(fs.readFileSync(path.join(__dirname,'../dist/word/taskpane.html'),'utf8').replace(/<script src="?https:\/\/appsforoffice[^>]+><\/script>/,mock)));
app.get('/api/me',(req,res)=>res.json({displayName:'预览用户',mail:'preview@example.com'}));
app.post('/api/translate',express.json(),(req,res)=>req.body.html.includes('模拟控件异常') ? res.status(409).json({error:'所选范围包含尚未恢复的翻译，请先恢复该翻译，再重新选择内容。'}) : res.json({html:'<div>Hello, Jason.</div>'}));
app.post('/preview-error',express.text(),(req,res)=>{console.error(req.body);res.sendStatus(204)});
app.use(express.static(path.join(__dirname,'../dist')));
app.listen(4173,'127.0.0.1',()=>console.log('Word UI preview http://localhost:4173/word/taskpane.html (mock Office host)'));
