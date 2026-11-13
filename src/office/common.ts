/** Công cụ dùng chung mọi ứng dụng: đọc kỹ năng (use_skill) và đọc chéo nội dung ứng dụng khác (read_other_app). */
import { ExtraTools, ToolSpec } from "../core/agent/types";
import { Skill, useSkill } from "../core/skills";
import { readOtherApp } from "../core/share";

const HOSTS = ["Excel", "Word", "PowerPoint"];

export function createCommonTools(o: { host: string; skills: Skill[]; shareContext: boolean }): ExtraTools {
  const specs: ToolSpec[] = [];
  if (o.skills.length) {
    specs.push({
      name: "use_skill",
      description:
        "Read the full instructions of a skill from the skills list in the system prompt (or one of its reference files). Call it before doing a task that matches the skill's description, then follow the instructions.",
      parameters: {
        type: "object",
        properties: {
          name: { type: "string", description: "Skill name exactly as listed." },
          file: { type: "string", description: "Optional reference file of the skill." },
        },
        required: ["name"],
      },
    });
  }
  if (o.shareContext) {
    specs.push({
      name: "read_other_app",
      description:
        "Read what is currently open in another Office app (Excel workbook, Word document or PowerPoint deck) that has the AI Đa Nguồn pane open — e.g. to use Excel figures in a Word report or in slides.",
      parameters: {
        type: "object",
        properties: { app: { type: "string", enum: HOSTS.filter((h) => h !== o.host) } },
        required: ["app"],
      },
    });
  }
  return {
    specs,
    execute: async (name, args) => {
      if (name === "use_skill") return useSkill(o.skills, args?.name, args?.file);
      if (name === "read_other_app") return readOtherApp(args?.app, o.host);
      throw new Error(`Không có công cụ ${name}`);
    },
  };
}
