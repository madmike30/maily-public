'use strict';
const fs=require('node:fs'),path=require('node:path'),crypto=require('node:crypto');
const {executeProcess,verifyRuntime}=require('../../scripts/connect-outlook.cjs');
(async()=>{
  const root=path.resolve(process.argv[2]);verifyRuntime(root);
  const home=path.resolve('output/cli-offline-check-'+crypto.randomUUID());fs.mkdirSync(home,{recursive:true});
  for(const args of [['--version','--telemetry','false'],['install','--help','--telemetry','false'],['auth','login','m365','--help','--telemetry','false']]) {
    const result=await executeProcess(path.join(root,'node.exe'),[path.join(root,'atk/cli.js'),...args],home,30000);
    if(result.code!==0)throw new Error('Offline CLI check failed');
    fs.writeFileSync(path.join(home,args.join('-').replace(/[^a-z0-9-]/gi,'')+'.log'),result.stdout);
    console.log('PASS pinned ATK '+args.join(' '));
  }
  const source=`const keytar=require(${JSON.stringify(path.join(root,'atk/node_modules/keytar'))});if(typeof keytar.getPassword!=='function'||typeof keytar.setPassword!=='function')process.exit(1);`;
  const native=await executeProcess(path.join(root,'node.exe'),['-e',source],home,15000);
  if(native.code!==0)throw new Error('Native keytar failed');
  console.log('PASS pinned native keytar module loads; no credential reads or writes');
})().catch(error=>{console.error(error.message);process.exitCode=1;});
