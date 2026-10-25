'use strict';
const path=require('node:path');
const crypto=require('node:crypto');
async function check(keytar) {
  const service='Microsoft 365 Agents Toolkit',account='m365';
  let key=await keytar.getPassword(service,account);
  if(key!==null&&key!==undefined&&key.length!==32) return false;
  if(key===null||key===undefined) {
    // ATK's documented implementation shares this encryption key, not account tokens.
    // Never replace an existing key; malformed storage requires manual repair.
    key=crypto.randomBytes(16).toString('hex');
    const second=await keytar.getPassword(service,account);
    if(second!==null&&second!==undefined) return second.length===32;
    await keytar.setPassword(service,account,key);
    return await keytar.getPassword(service,account)===key;
  }
  return true;
}
async function main() {
  try {
    const root=path.resolve(process.argv[2]),atk=path.join(root,'atk');
    // This deployment uses ATK's browser fallback, not a shared Windows broker.
    try {require.resolve('@azure/msal-node-extensions',{paths:[path.join(atk,'lib')]});return 1;} catch(error) {if(error.code!=='MODULE_NOT_FOUND')return 1;}
    const keytarPath=require.resolve('keytar',{paths:[path.join(atk,'lib')]});
    if(!keytarPath.startsWith(path.join(atk,'node_modules','keytar')+path.sep)) return 1;
    return await check(require(keytarPath))?0:1;
  } catch (_) {return 1;}
}
module.exports={check};
if(require.main===module) main().then(code=>{process.exitCode=code;});
