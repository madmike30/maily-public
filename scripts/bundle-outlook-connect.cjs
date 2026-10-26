'use strict';
const fs=require('node:fs'),path=require('node:path'),crypto=require('node:crypto'),zlib=require('node:zlib');
const hash=bytes=>crypto.createHash('sha256').update(bytes).digest('hex');
function extract(bytes,destination,stripPackage,{omitSourceMaps=false}={}) {
  const tar=zlib.gunzipSync(bytes);
  for(let offset=0;offset+512<=tar.length;) {
    const name=tar.toString('utf8',offset,offset+100).replace(/\0.*$/s,'');
    if(!name)break;
    const prefix=tar.toString('utf8',offset+345,offset+500).replace(/\0.*$/s,'');
    const raw=prefix?prefix+'/'+name:name;
    const size=parseInt(tar.toString('ascii',offset+124,offset+136).replace(/\0.*$/s,'').trim()||'0',8);
    const type=tar.toString('ascii',offset+156,offset+157);
    if(!Number.isSafeInteger(size)||size<0||offset+512+size>tar.length)throw new Error('Invalid archive');
    if(type==='x'||type==='g') {offset+=512+Math.ceil(size/512)*512;continue;}
    const relative=stripPackage?raw.replace(/^package\//,''):raw;
    if(path.isAbsolute(relative)||relative.split(/[\\/]/).includes('..')||relative.includes(':'))throw new Error('Unsafe archive path');
    const target=path.resolve(destination,relative);
    if(!target.startsWith(path.resolve(destination)+path.sep)&&target!==path.resolve(destination))throw new Error('Archive escapes destination');
    if(type==='5')fs.mkdirSync(target,{recursive:true});
    else if(type==='0'||type==='\0') {if(!(omitSourceMaps&&relative.endsWith('.js.map'))){fs.mkdirSync(path.dirname(target),{recursive:true});fs.writeFileSync(target,tar.subarray(offset+512,offset+512+size),{flag:'wx'});}}
    else throw new Error('Archive links and unsupported entries are forbidden');
    offset+=512+Math.ceil(size/512)*512;
  }
}
async function bundle(destination,cache) {
  const lock=JSON.parse(fs.readFileSync(path.join(__dirname,'outlook-connect-runtime.lock.json'),'utf8'));
  if(fs.existsSync(destination))throw new Error('Choose a new output directory; evidence is never overwritten');
  fs.mkdirSync(destination,{recursive:true});fs.mkdirSync(cache,{recursive:true});
  for(const file of lock.files) {
    if(path.basename(file.name)!==file.name)throw new Error('Unsafe locked path');
    const cached=path.join(cache,file.name);let bytes;
    if(fs.existsSync(cached))bytes=fs.readFileSync(cached);
    else {const response=await fetch(file.url,{signal:AbortSignal.timeout(120000)});if(!response.ok)throw new Error('Locked runtime download failed');bytes=Buffer.from(await response.arrayBuffer());if(hash(bytes)!==file.sha256)throw new Error('Runtime download checksum mismatch');fs.writeFileSync(cached,bytes,{flag:'wx'});}
    if(hash(bytes)!==file.sha256)throw new Error('Cached runtime checksum mismatch; preserve it and use a new cache');
    if(file.npmPath) {
      if(!/^atk\/node_modules\/[a-z0-9-]+$/.test(file.npmPath))throw new Error('Unsafe dependency destination');
      extract(bytes,path.join(destination,file.npmPath),true);
    }
    else if(file.name==='atk.tgz')extract(bytes,path.join(destination,'atk'),true,{omitSourceMaps:true});
    else if(file.name==='keytar.tgz')extract(bytes,path.join(destination,'atk/node_modules/keytar'),true);
    else if(file.name==='keytar-native.tgz')extract(bytes,path.join(destination,'atk/node_modules/keytar'),false);
    else fs.writeFileSync(path.join(destination,file.name),bytes,{flag:'wx'});
  }
  const files=[];
  const visit=dir=>{for(const entry of fs.readdirSync(dir,{withFileTypes:true})){const full=path.join(dir,entry.name);if(entry.isDirectory())visit(full);else if(entry.isFile())files.push({path:path.relative(destination,full).split(path.sep).join('/'),sha256:hash(fs.readFileSync(full))});else throw new Error('Non-file runtime entry');}};
  visit(destination);
  fs.writeFileSync(path.join(destination,'runtime-inventory.json'),JSON.stringify({version:lock.version,nodeVersion:lock.nodeVersion,files},null,2)+'\n');
  fs.copyFileSync(path.join(__dirname,'outlook-connect-runtime.lock.json'),path.join(destination,'outlook-connect-runtime.lock.json'));
  console.log(`Verified pinned Outlook connector: ${files.length} runtime files.`);
}
module.exports={extract,bundle};
if(require.main===module)bundle(path.resolve(process.argv[2]),path.resolve(process.argv[3]||'output/outlook-connect-downloads')).catch(error=>{console.error(error.message);process.exitCode=1;});
