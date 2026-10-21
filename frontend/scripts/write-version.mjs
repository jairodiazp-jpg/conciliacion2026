import { execSync } from "node:child_process";
import { mkdirSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const outputFile = path.resolve(scriptDir, "../src/generated/buildInfo.js");

function runCommand(command) {
  try {
    return execSync(command, { encoding: "utf8", stdio: ["ignore", "pipe", "ignore"] }).trim();
  } catch {
    return "";
  }
}

function buildVersion() {
  const commitHash = (
    process.env.RENDER_GIT_COMMIT ||
    process.env.SOURCE_VERSION ||
    process.env.GIT_COMMIT ||
    process.env.HEROKU_SLUG_COMMIT ||
    process.env.COMMIT_REF ||
    runCommand("git rev-parse --short HEAD")
  ).trim();

  const commitCount = runCommand("git rev-list --count HEAD");

  return "2.0";
}

const version = buildVersion();
const buildDate = new Date().toISOString();

mkdirSync(path.dirname(outputFile), { recursive: true });

writeFileSync(
  outputFile,
  `export const APP_VERSION = ${JSON.stringify(version)};\nexport const APP_BUILD_DATE = ${JSON.stringify(buildDate)};\n`,
  "utf8"
);

console.log(`Generated ${outputFile} (${version})`);