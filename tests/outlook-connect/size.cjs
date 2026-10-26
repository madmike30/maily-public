'use strict';
const assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path'),zlib=require('node:zlib');
const {extract}=require('../../scripts/bundle-outlook-connect.cjs');
const root=path.resolve('output/connector-size-'+require('node:crypto').randomUUID());
function entry(name,body){const bytes=Buffer.from(body);const h=Buffer.alloc(512);h.write('package/'+name);h.write(bytes.length.toString(8).padStart(11,'0'),124);h.write('0',156);return Buffer.concat([h,bytes,Buffer.alloc((512-bytes.length%512)%512)]);}
const archive=zlib.gzipSync(Buffer.concat([entry('lib/index.js','module.exports=42;'),entry('lib/index.js.map','synthetic debug map'),entry('LICENSE','synthetic license'),Buffer.alloc(1024)]));
extract(archive,root,true,{omitSourceMaps:true});
assert.equal(fs.existsSync(path.join(root,'lib/index.js.map')),false,'development source maps must be omitted');
assert.equal(require(path.join(root,'lib/index.js')),42);
assert.equal(fs.readFileSync(path.join(root,'LICENSE'),'utf8'),'synthetic license');
const full=root+'-full';extract(archive,full,true);assert.equal(fs.existsSync(path.join(full,'lib/index.js.map')),true);
console.log('PASS compact extraction excludes only source maps; executable code/licenses and default extraction preserved');