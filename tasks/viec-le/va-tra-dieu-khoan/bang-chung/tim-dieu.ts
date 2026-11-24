import * as fs from "fs";
const JSZip = require("D:/ADD IN/node_modules/jszip");
import { docxParagraphs } from "D:/ADD IN/.claude/worktrees/multi-agent-orchestration-eval-53eeac/src/core/docxparas";
import { buildOutline, findClause } from "D:/ADD IN/.claude/worktrees/multi-agent-orchestration-eval-53eeac/src/core/outline";
(async () => { const z = await JSZip.loadAsync(fs.readFileSync(process.argv[2])); const g = async (n: string) => (await z.file(n)?.async("string")) || undefined;
  const rows = docxParagraphs({ document: (await g("word/document.xml"))!, styles: await g("word/styles.xml"), numbering: await g("word/numbering.xml") });
  const o = buildOutline(rows);
  for (const q of process.argv.slice(3)) console.log(JSON.stringify(q), "→", findClause(o, q).map((x: any) => `${x.num}@¶${x.i}-${x.end} lvl${x.level}`).join(" ; ") || "KHÔNG THẤY");
})();
