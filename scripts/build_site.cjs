const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const output = path.join(root, "public");
const data = path.join(output, "data");

for (const [directory, allowed] of [
  [output, ["index.html", "data"]],
  [data, ["carti.json"]],
]) {
  if (fs.existsSync(directory)) {
    const unexpected = fs.readdirSync(directory).filter(name => !allowed.includes(name));
    if (unexpected.length) {
      throw new Error(`Fișiere neașteptate în ${directory}: ${unexpected.join(", ")}`);
    }
  }
}

fs.mkdirSync(data, { recursive: true });
fs.copyFileSync(path.join(root, "biblioteca.html"), path.join(output, "index.html"));
fs.copyFileSync(path.join(root, "biblioteca.html"), path.join(root, "index.html"));
fs.copyFileSync(path.join(root, "data", "carti.json"), path.join(data, "carti.json"));
console.log("Site-ul public a fost pregătit in public/.");
