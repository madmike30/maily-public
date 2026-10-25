import fs from 'node:fs';
import cp from 'node:child_process';
import assert from 'node:assert/strict';
const path='services/Maily.OutlookAddin.Web/Program.cs';
const source=process.argv.includes('--baseline') ? cp.execFileSync('git',['show',`main:${path}`],{encoding:'utf8'}) : fs.readFileSync(path,'utf8');
const logs=source.match(/logger\.Log(?:Information|Warning|Error)\([\s\S]*?\);/g) || [];
assert.ok(logs.length>10);
for(const log of logs) {
 assert.doesNotMatch(log,/request\.(?:Client\.)?(?:MailboxAddress|Subject)|resolvedMailbox|DescribeRankedSubjects|entry\.(?:Detail|MailboxAddress)|LogError\(\s*ex,/,'sensitive source reaches logging sink');
}
const requestLog=source.slice(source.indexOf('static async Task AppendRequestLogAsync'),source.indexOf('static async Task AppendRequestLogAsync')+1000);
assert.doesNotMatch(requestLog,/Request\.QueryString|exception\?\.Message/);
assert.doesNotMatch(source,/error = ex\.Message|StyleLastRefreshError = ex\.Message/);
console.log('PASS backend diagnostic privacy sink contracts');

const readiness=source.slice(source.indexOf('static (bool Ok, object Details, string UserMessage, string RecommendedAction) BuildCodexReadiness'), source.indexOf('static CodexProcessResult RunCodexProcess'));
assert.doesNotMatch(readiness,/codexVersionOutput\s*=|loginStatusOutput\s*,/,'readiness must never return raw provider output');
console.log('PASS readiness privacy sink contract');
const packaging=fs.readFileSync('scripts/build-maily-outlook-installer.ps1','utf8');
const setupExecutable=packaging.match(/\$setupExeName\s*=\s*"([^"]+)"/)[1];
const readinessClassifier=fs.readFileSync('src/Maily.Generation/Services/CodexReadiness.cs','utf8');
assert.ok(readinessClassifier.includes(`Open ${setupExecutable} from the original extracted installer package and`),'readiness recovery must name the packaged installer');
console.log('PASS readiness packaged installer guidance');
