/* Maily's supported Microsoft Toolkit adapter. Never forward provider output. */
'use strict';
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const {spawn,execFile} = require('node:child_process');
const ADDIN_ID = '683ea353-6d71-4315-afd4-147e968605a9';
const failed = (code, message) => ({status:'failed',code,message});

async function connect(options, deps) {
  try {
    await deps.verify();
    const home = await deps.createSession();
    const preflight = await deps.execute(['credential-preflight'], home);
    if (preflight.code !== 0) return failed('credential-storage', 'Microsoft sign-in storage is unavailable. No account was connected.');
    const login = await deps.execute(['auth','login','m365','--telemetry','false'], home);
    if (login.code !== 0) return failed(login.timedOut ? 'sign-in-timeout' : 'sign-in-incomplete', 'Microsoft sign-in was not completed. Your existing installation is unchanged.');
    const install = await deps.execute(['install','--xml-path',options.manifest,'--interactive','false','--telemetry','false'], home);
    if (install.code !== 0) return failed(install.timedOut ? 'registration-timeout' : 'registration-failed', 'Outlook connection was not confirmed. Check Microsoft account permissions, then check the connection before retrying.');
    const appId = /(?:^|\n)\s*(?:\[[^\]\r\n]*\]\s*)?AppId:\s*([a-f0-9-]{36})\s*(?:\r?\n|$)/i.exec(install.stdout || '')?.[1];
    const titleId = /(?:^|\n)\s*(?:\[[^\]\r\n]*\]\s*)?TitleId:\s*([A-Za-z0-9_-]{1,160})\s*(?:\r?\n|$)/i.exec(install.stdout || '')?.[1];
    if (appId?.toLowerCase() !== options.manifestId.toLowerCase() || !titleId) return failed('unverified-result', 'Microsoft returned an unverified connection result. Check Outlook before retrying.');
    return {status:'registered',code:'registered',message:'Microsoft accepted Maily. Open Outlook and check the connection.',manifestId:options.manifestId,titleId};
  } catch (_) {
    return failed('connection-unavailable', 'The Outlook connection component could not run. Repair Maily and try again.');
  }
}

function isolatedEnvironment(home) {
  const env = {...process.env,USERPROFILE:home,HOME:home,APPDATA:path.join(home,'AppData','Roaming'),LOCALAPPDATA:path.join(home,'AppData','Local'),TEMP:path.join(home,'Temp'),TMP:path.join(home,'Temp'),NO_COLOR:'1',FORCE_COLOR:'0'};
  for (const key of Object.keys(env)) if (/^(NODE_OPTIONS|NODE_PATH|AZURE_|MSAL_|SIDELOADING_|TEAMSFX_|ATK_)/i.test(key)) delete env[key];
  return env;
}

function executeProcess(executable,args,home,timeoutMs=300000) {
  return new Promise(resolve=>{
    const env=isolatedEnvironment(home);
    for(const folder of [env.APPDATA,env.LOCALAPPDATA,env.TEMP]) fs.mkdirSync(folder,{recursive:true});
    const child=spawn(executable,args,{cwd:home,env,windowsHide:true,stdio:['ignore','pipe','pipe']});
    let stdout='',tooLarge=false,timedOut=false;
    const limit=128*1024;
    let stopping=false;
    const stop=()=>{
      if(stopping)return;stopping=true;
      if(process.platform==='win32') {
        // Only terminate this adapter-owned CLI process and descendants.
        execFile(path.join(process.env.SystemRoot || 'C:\\Windows','System32','taskkill.exe'),['/PID',String(child.pid),'/T','/F'],{windowsHide:true,timeout:2500,maxBuffer:4096},()=>{if(child.exitCode===null)child.kill();});
      } else child.kill('SIGKILL');
    };
    const timer=setTimeout(()=>{timedOut=true;stop();},timeoutMs);
    let bytes=0;
    child.stdout.on('data',data=>{bytes+=data.length;if(bytes>limit){tooLarge=true;stop();}else stdout+=data.toString('utf8');});
    child.stderr.on('data',data=>{bytes+=data.length;if(bytes>limit){tooLarge=true;stop();}});
    child.on('error',()=>{clearTimeout(timer);resolve({code:1,stdout:''});});
    child.on('close',code=>{clearTimeout(timer);resolve({code:tooLarge||timedOut?1:code,stdout:tooLarge?'':stdout,timedOut});});
  });
}

function verifyRuntime(root) {
  const inventory=JSON.parse(fs.readFileSync(path.join(root,'runtime-inventory.json'),'utf8'));
  if(inventory.version!=='1.1.18'||inventory.nodeVersion!=='v22.23.4'||!Array.isArray(inventory.files)||inventory.files.length<10) throw new Error('Invalid runtime');
  for(const item of inventory.files) {
    const file=path.resolve(root,item.path);
    if(!file.startsWith(path.resolve(root)+path.sep)||fs.lstatSync(file).isSymbolicLink()) throw new Error('Unsafe inventory');
    if(crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex')!==item.sha256) throw new Error('Runtime hash mismatch');
  }
  for(const file of ['node.exe','atk/cli.js','atk/node_modules/keytar/lib/keytar.js','atk/node_modules/keytar/build/Release/keytar.node']) if(!inventory.files.some(item=>item.path===file)) throw new Error('Missing runtime file');
  const allowed=new Set(inventory.files.map(item=>item.path));
  allowed.add('runtime-inventory.json');allowed.add('outlook-connect-runtime.lock.json');
  const visit=dir=>{for(const entry of fs.readdirSync(dir,{withFileTypes:true})){const full=path.join(dir,entry.name);if(entry.isDirectory())visit(full);else if(!entry.isFile()||!allowed.has(path.relative(root,full).split(path.sep).join('/')))throw new Error('Unlisted runtime entry');}};
  visit(root);
}

async function main(argv) {
  const read=name=>{const index=argv.indexOf(name);if(index<0||!argv[index+1])throw new Error('Missing option');return path.resolve(argv[index+1]);};
  try {
    const root=read('--runtime-root'), manifest=read('--manifest'), sessionRoot=read('--session-root');
    const xml=fs.readFileSync(manifest,'utf8');
    if(!new RegExp('<Id>\\s*'+ADDIN_ID+'\\s*</Id>','i').test(xml)) throw new Error('Wrong manifest');
    const result=await connect({manifest,manifestId:ADDIN_ID},{
      verify:async()=>verifyRuntime(root),
      createSession:async()=>{fs.mkdirSync(sessionRoot,{recursive:true});return fs.mkdtempSync(path.join(sessionRoot,'connect-'));},
      execute:(args,home)=>executeProcess(path.join(root,'node.exe'),args[0]==='credential-preflight'?[path.join(__dirname,'outlook-credential-preflight.cjs'),root]:[path.join(root,'atk','cli.js'),...args],home,args[0]==='credential-preflight'?15000:300000)
    });
    process.stdout.write(JSON.stringify(result)+'\n');
    process.exitCode=result.status==='registered'?0:1;
  } catch (_) {process.stdout.write(JSON.stringify(failed('invalid-package','The Outlook connection package is incomplete. Repair Maily and try again.'))+'\n');process.exitCode=1;}
}
module.exports={connect,isolatedEnvironment,executeProcess,verifyRuntime};
if(require.main===module) void main(process.argv.slice(2));
