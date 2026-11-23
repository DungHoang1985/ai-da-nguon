const fs=require("fs");const S=process.argv[2];
for (const [d,arm] of [["bo-1","A"],["bo-1","B"],["va","B"]]) for (const f of fs.readdirSync(S+"/"+d).filter(f=>(f.startsWith("H1-"+arm)||f.startsWith("R1-"+arm))&&f.endsWith(".json"))) {
 const k=JSON.parse(fs.readFileSync(S+"/"+d+"/"+f,"utf8")).kq.khung||""; const steps=k.split("\n").filter(l=>l.startsWith("✅ Đọc")).map(l=>l.slice(2)).join(", ");
 console.log((d=="va"?"BẢN VÁ ":"4.11.8 "+(arm=="B"?"(gọn) ":"")).padEnd(14), f.padEnd(12), steps);}
