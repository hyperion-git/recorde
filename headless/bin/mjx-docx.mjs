#!/usr/bin/env node
import { main } from '../cli.mjs';
main(process.argv.slice(2)).catch((e) => { process.stderr.write(`mjx-docx: ${e.message}\n`); process.exit(1); });
