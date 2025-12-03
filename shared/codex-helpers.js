#!/usr/bin/env node

// Cross-platform Codex helper fallback (PowerShell-free).
// Mirrors shared/codex-helpers.ps1 behaviors for context resolution, locking,
// JSON write/validation, and basic /save and /end flows.

import fs from 'fs';
import { promises as fsp } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { execFileSync } from 'child_process';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const DEFAULT_LOCK_TIMEOUT_MS = 30_000;
const DEFAULT_LOCK_RETRY_MS = 250;
const DEFAULT_RETENTION_KEEP = 3;
const RETENTION_CONFIG_PATH = path.join(__dirname, '../prompts/config/checkpoints.retention.json');

async function loadJsonSafe(filePath, fallback) {
  try {
    const raw = await fsp.readFile(filePath, 'utf8');
    return JSON.parse(raw);
  } catch {
    return fallback;
  }
}

async function getRetentionKeepCount(override) {
  if (typeof override === 'number' && override > 0) {
    return override;
  }
  const cfg = await loadJsonSafe(RETENTION_CONFIG_PATH, {});
  const keep = Number.parseInt(cfg?.keep_latest, 10);
  if (Number.isInteger(keep) && keep > 0) {
    return keep;
  }
  return DEFAULT_RETENTION_KEEP;
}

async function ensureDirectory(p) {
  const dir = path.dirname(p);
  if (!dir) return;
  await fsp.mkdir(dir, { recursive: true });
}

async function writeJson(filePath, inputObject) {
  await ensureDirectory(filePath);
  const json = JSON.stringify(inputObject, null, 2);
  await fsp.writeFile(filePath, json, { encoding: 'utf8' });
}

async function writeMarkdown(filePath, content) {
  await ensureDirectory(filePath);
  await fsp.writeFile(filePath, content, { encoding: 'utf8' });
}

function matchesType(value, expected) {
  switch (expected) {
    case 'object':
      return value !== null && typeof value === 'object' && !Array.isArray(value);
    case 'array':
      return Array.isArray(value);
    case 'string':
      return typeof value === 'string';
    case 'boolean':
      return typeof value === 'boolean';
    case 'number':
      return typeof value === 'number' && Number.isFinite(value);
    case 'integer':
      return Number.isInteger(value);
    case 'null':
      return value === null;
    default:
      return true;
  }
}

function validateSchemaValue(value, schema, pathKey, errors) {
  if (!schema) return;

  const typeSpec = schema.type;
  const expectedTypes = Array.isArray(typeSpec) ? typeSpec : typeSpec ? [typeSpec] : [];
  let resolvedType = null;

  if (expectedTypes.length > 0) {
    const matched = expectedTypes.find((type) => matchesType(value, type));
    if (!matched) {
      errors.push(
        `${pathKey}: expected type '${expectedTypes.join(', ')}', got '${
          value === null ? 'null' : Array.isArray(value) ? 'array' : typeof value
        }'.`
      );
      return;
    }
    resolvedType = matched;
  }

  if (schema.enum) {
    const matchesEnum = schema.enum.some((option) => Object.is(option, value));
    if (!matchesEnum) {
      errors.push(`${pathKey}: value '${value}' is not in enum [${schema.enum.join(', ')}]`);
    }
  }

  switch (resolvedType) {
    case 'object':
      validateObjectSchema(value, schema, pathKey, errors);
      break;
    case 'array':
      validateArraySchema(value, schema, pathKey, errors);
      break;
    default:
      break;
  }
}

function validateObjectSchema(value, schema, pathKey, errors) {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    errors.push(`${pathKey}: expected object but got '${typeof value === 'object' ? 'array' : typeof value}'.`);
    return;
  }

  if (schema.required) {
    schema.required.forEach((key) => {
      if (!Object.prototype.hasOwnProperty.call(value, key)) {
        errors.push(`${pathKey}: missing required property '${key}'.`);
      }
    });
  }

  if (schema.additionalProperties === false) {
    const allowed = schema.properties ? Object.keys(schema.properties) : [];
    Object.keys(value).forEach((key) => {
      if (allowed.length === 0 || !allowed.includes(key)) {
        errors.push(`${pathKey}: unexpected property '${key}'.`);
      }
    });
  }

  if (schema.properties) {
    Object.entries(schema.properties).forEach(([key, childSchema]) => {
      if (Object.prototype.hasOwnProperty.call(value, key)) {
        validateSchemaValue(value[key], childSchema, `${pathKey}.${key}`, errors);
      }
    });
  }
}

function validateArraySchema(value, schema, pathKey, errors) {
  if (!Array.isArray(value)) {
    errors.push(`${pathKey}: expected array but got '${typeof value}'.`);
    return;
  }

  if (schema.minItems !== undefined && value.length < schema.minItems) {
    errors.push(`${pathKey}: expected at least ${schema.minItems} items but found ${value.length}.`);
  }

  if (schema.maxItems !== undefined && value.length > schema.maxItems) {
    errors.push(`${pathKey}: expected at most ${schema.maxItems} items but found ${value.length}.`);
  }

  if (schema.items) {
    value.forEach((item, index) => {
      validateSchemaValue(item, schema.items, `${pathKey}[${index}]`, errors);
    });
  }
}

async function validateJson(schemaPath, jsonPath) {
  return validateJsonWithOptions(schemaPath, jsonPath, {});
}

async function validateJsonWithOptions(schemaPath, jsonPath, options = {}) {
  if (!fs.existsSync(schemaPath)) {
    throw new Error(`Schema file not found: ${schemaPath}`);
  }
  if (!fs.existsSync(jsonPath)) {
    throw new Error(`JSON file not found: ${jsonPath}`);
  }

  const schemaObject = JSON.parse(await fsp.readFile(schemaPath, 'utf8'));
  const jsonObject = JSON.parse(await fsp.readFile(jsonPath, 'utf8'));

  const errors = [];
  validateSchemaValue(jsonObject, schemaObject, '$', errors);

  if (errors.length > 0) {
    const message = `JSON at ${jsonPath} failed schema validation:\n${errors.join('\n')}`;
    if (options.warnOnly) {
      return {
        path: jsonPath,
        schema: schemaPath,
        valid: false,
        errors,
        checked_at: new Date().toISOString(),
      };
    }
    throw new Error(message);
  }

  return {
    path: jsonPath,
    schema: schemaPath,
    valid: true,
    checked_at: new Date().toISOString(),
  };
}

async function useLock(lockPath, fn, options = {}) {
  const timeoutMs = (options.timeoutSeconds || 0) * 1000 || DEFAULT_LOCK_TIMEOUT_MS;
  const retryMs = options.retryDelayMilliseconds || DEFAULT_LOCK_RETRY_MS;

  await ensureDirectory(lockPath);
  const deadline = Date.now() + timeoutMs;
  let handle;

  while (!handle) {
    try {
      handle = fs.openSync(lockPath, fs.constants.O_CREAT | fs.constants.O_EXCL | fs.constants.O_RDWR);
    } catch (error) {
      if (Date.now() > deadline) {
        throw new Error(`Unable to acquire lock at ${lockPath} within ${timeoutMs / 1000} seconds.`);
      }
      await new Promise((resolve) => setTimeout(resolve, retryMs));
    }
  }

  try {
    const payload = `locked ${new Date().toISOString()}\n`;
    fs.writeSync(handle, payload);
    fs.fsyncSync(handle);
    await fn();
  } finally {
    if (handle !== undefined) {
      fs.closeSync(handle);
    }
    try {
      await fsp.unlink(lockPath);
    } catch {
      // Ignore cleanup errors
    }
  }
}

function formatOutput({ summary, details = [], nextSteps = [], status = 'OK' }) {
  const lines = [];
  lines.push('Summary');
  lines.push(summary);

  if (details.length) {
    lines.push('');
    lines.push('Details');
    details.forEach((line) => lines.push(`- ${line}`));
  }

  if (nextSteps.length) {
    lines.push('');
    lines.push('Next Steps');
    nextSteps.forEach((line) => lines.push(`- ${line}`));
  }

  lines.push('');
  lines.push(`STATUS: ${status}`);
  return lines.join('\n');
}

function getProjectContext(projectRoot = process.cwd(), referenceTime = new Date()) {
  const pad = (value) => String(value).padStart(2, '0');
  const todayKey = `${pad(referenceTime.getUTCDate())}-${pad(referenceTime.getUTCMonth() + 1)}-${String(
    referenceTime.getUTCFullYear()
  ).slice(-2)}`;

  const dailyDir = path.join(projectRoot, 'docs', todayKey);
  const locksDir = path.join(dailyDir, 'locks');

  return {
    projectRoot,
    todayKey,
    dailyDir,
    sessionPath: path.join(dailyDir, '.session'),
    locksDir,
  };
}

function listDirNames(dirPath) {
  try {
    return fs
      .readdirSync(dirPath, { withFileTypes: true })
      .filter((entry) => entry.isDirectory())
      .map((entry) => entry.name);
  } catch {
    return [];
  }
}

function getOpenSpecStats(projectRoot) {
  const openspecRoot = path.join(projectRoot, 'openspec');
  if (!fs.existsSync(openspecRoot)) {
    return { hasOpenSpec: false };
  }

  const changes = listDirNames(path.join(openspecRoot, 'changes'));
  const specs = listDirNames(path.join(openspecRoot, 'specs'));
  const pending = changes.filter((name) =>
    fs.existsSync(path.join(openspecRoot, 'changes', name, 'proposal.md'))
  );

  return {
    hasOpenSpec: true,
    changeCount: changes.length,
    specCount: specs.length,
    pendingChangeCount: pending.length,
  };
}

function getActivityStats(dailyDir) {
  const workLogPath = path.join(dailyDir, 'logs', 'work.json');
  if (!fs.existsSync(workLogPath)) {
    return { workSessions: 0, activeTasks: 0, lastIteration: null };
  }

  try {
    const content = JSON.parse(fs.readFileSync(workLogPath, 'utf8'));
    const sessions = content.work_sessions || [];
    const activeTasks = sessions.filter((session) => session.status === 'in_progress').length;
    const lastIteration = sessions
      .map((session) => session.continue_iterations)
      .filter(Boolean)
      .flat()
      .sort((a, b) => (a.timestamp || '').localeCompare(b.timestamp || ''))
      .pop();

    return {
      workSessions: sessions.length,
      activeTasks,
      lastIteration: lastIteration ? lastIteration.timestamp : null,
    };
  } catch {
    return { workSessions: 0, activeTasks: 0, lastIteration: null };
  }
}

function getCoachingTip() {
  const tips = [
    'Ship in vertical slices: demo something end-to-end each day.',
    'Prefer measurable outcomes over busywork; attach tests or metrics to every change.',
    'Document decisions inline while they are fresh to avoid context drift.',
    'Keep feedback loops tight: run smoke tests before deep refactors.',
    'Simplify tooling when stuck; remove one variable before adding another.',
  ];
  return tips[Math.floor(Math.random() * tips.length)];
}

async function appendOpsLog(context, entry) {
  try {
    const logPath = path.join(context.locksDir, 'ops.log');
    await ensureDirectory(logPath);
    const line = `${JSON.stringify({ ...entry, ts: new Date().toISOString() })}\n`;
    await fsp.appendFile(logPath, line, { encoding: 'utf8' });
  } catch {
    // Best-effort only
  }
}

function execGit(args, projectRoot) {
  try {
    return execFileSync('git', args, {
      cwd: projectRoot,
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'ignore'],
    }).trim();
  } catch {
    return null;
  }
}

function collectGitInfo(projectRoot, options = {}) {
  const includeDiffstat = options.includeDiffstat !== false;
  const branch = execGit(['rev-parse', '--abbrev-ref', 'HEAD'], projectRoot) || 'unknown';
  const head = execGit(['rev-parse', '--short', 'HEAD'], projectRoot) || 'unknown';
  const statusOutput = execGit(['status', '--porcelain'], projectRoot) || '';
  const dirty = statusOutput.length > 0;
  const diffstat = includeDiffstat
    ? execGit(['diff', '--stat'], projectRoot) || (dirty ? 'dirty' : 'clean')
    : dirty
    ? 'dirty'
    : 'clean';

  const filesChanged = statusOutput
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => line.replace(/^[A-Z?]{1,2}\s+/, ''));

  return {
    git: {
      branch,
      head,
      dirty,
      diffstat,
    },
    filesChanged,
  };
}

function getCheckpointId(reference = new Date()) {
  const pad = (num) => String(num).padStart(2, '0');
  return `${reference.getUTCFullYear()}${pad(reference.getUTCMonth() + 1)}${pad(reference.getUTCDate())}-${pad(
    reference.getUTCHours()
  )}${pad(reference.getUTCMinutes())}${pad(reference.getUTCSeconds())}`;
}

function loadSession(sessionPath) {
  if (!fs.existsSync(sessionPath)) {
    return null;
  }

  try {
    return JSON.parse(fs.readFileSync(sessionPath, 'utf8'));
  } catch {
    return null;
  }
}

async function writeSession(sessionPath, payload) {
  const stored = { ...payload, project_root: '{PROJECT_ROOT}' };
  await writeJson(sessionPath, stored);
  await validateJsonWithOptions(path.join(__dirname, '../prompts/config/session.schema.json'), sessionPath, {
    warnOnly: true,
  });
}

async function saveCheckpoint(options) {
  const now = options.referenceTime || new Date();
  const context = getProjectContext(options.projectRoot, now);
  const { git, filesChanged } = collectGitInfo(context.projectRoot, {
    includeDiffstat: options.includeDiffstat !== false,
  });
  const id = options.id || getCheckpointId(now);
  const checkpointPath = path.join(context.dailyDir, 'checkpoints', `${id}.json`);
  const checkpointsDir = path.dirname(checkpointPath);

  const payload = {
    id,
    status: options.status || 'open',
    created_at: now.toISOString(),
    session_status: options.sessionStatus || 'ACTIVE',
    git,
    files_changed: filesChanged,
    tests: {
      ran: Boolean(options.testsRan),
      summary: options.testsSummary || (options.testsRan ? 'Completed' : 'Not run'),
      ...(options.testsDetails ? { details: options.testsDetails } : {}),
      ...(options.testsRecommended ? { recommended: options.testsRecommended } : {}),
    },
    tasks: {
      in_progress: options.tasksInProgress || [],
      completed: options.tasksCompleted || [],
    },
    notes: options.notes || '',
    created_by: options.createdBy || '/save',
  };

  if (options.alerts && options.alerts.length > 0) {
    payload.alerts = options.alerts;
  }
  if (options.blockers && options.blockers.length > 0) {
    payload.blockers = options.blockers;
  }

  await ensureDirectory(checkpointPath);

  if (options.dryRun) {
    return { checkpointPath, payload, dryRun: true, retentionDeleted: [] };
  }

  const retentionDeleted = [];
  await useLock(path.join(context.locksDir, 'checkpoint.lock'), async () => {
    await writeJson(checkpointPath, payload);
    await validateJson(path.join(__dirname, '../prompts/config/checkpoints.schema.json'), checkpointPath);
    const keepLatest = await getRetentionKeepCount(options.keepLatest);
    const deletions = await enforceCheckpointRetention(checkpointsDir, keepLatest);
    retentionDeleted.push(...deletions);
    await appendOpsLog(context, {
      op: 'checkpoint',
      id,
      path: checkpointPath,
      keep_latest: keepLatest,
      retention_deleted: deletions,
    });
  });

  return { checkpointPath, payload, retentionDeleted };
}

async function enforceCheckpointRetention(checkpointsDir, keepLatest = 3) {
  await ensureDirectory(path.join(checkpointsDir, 'retention.json'));

  let entries = [];
  try {
    entries = await fsp.readdir(checkpointsDir, { withFileTypes: true });
  } catch {
    return [];
  }

  const candidates = [];
  for (const entry of entries) {
    if (!entry.isFile()) continue;
    if (!entry.name.endsWith('.json')) continue;
    if (entry.name === 'retention.json') continue;
    const id = path.basename(entry.name, '.json');
    candidates.push({ id, name: entry.name });
  }

  candidates.sort((a, b) => b.name.localeCompare(a.name));
  const toDelete = candidates.slice(keepLatest);
  if (toDelete.length === 0) {
    return [];
  }

  const retentionPath = path.join(checkpointsDir, 'retention.json');
  let retentionLog = [];
  try {
    const existing = JSON.parse(await fsp.readFile(retentionPath, 'utf8'));
    if (Array.isArray(existing)) {
      retentionLog = existing;
    }
  } catch {
    retentionLog = [];
  }

  for (const entry of toDelete) {
    const target = path.join(checkpointsDir, entry.name);
    try {
      await fsp.unlink(target);
      retentionLog.push({
        timestamp: new Date().toISOString(),
        deleted_id: entry.id,
      });
    } catch {
      // Ignore deletion failures; keep best-effort retention
    }
  }

  await writeJson(retentionPath, retentionLog);
  return toDelete.map((d) => d.id);
}

async function endSession(options) {
  const now = options.referenceTime || new Date();
  const context = getProjectContext(options.projectRoot, now);
  const session = loadSession(context.sessionPath);

  const payload = session
    ? {
        ...session,
        status: options.sessionStatus || 'CLOSED',
        last_activity: now.toISOString(),
        last_command: 'end',
      }
    : {
        session_id: `${context.todayKey.replace(/-/g, '')}_${getCheckpointId(now).split('-')[1]}`,
        session_started: now.toISOString(),
        date: context.todayKey,
        project_root: '{PROJECT_ROOT}',
        status: options.sessionStatus || 'CLOSED',
        last_activity: now.toISOString(),
        last_command: 'end',
      };

  if (options.dryRun) {
    return { sessionPath: context.sessionPath, payload, dryRun: true };
  }

  await useLock(path.join(context.locksDir, 'session.lock'), async () => {
    await writeSession(context.sessionPath, payload);
    await appendOpsLog(context, {
      op: 'session_end',
      path: context.sessionPath,
      status: payload.status,
      last_command: payload.last_command,
    });
  });

  return { sessionPath: context.sessionPath, payload };
}

async function heartbeatSession(options) {
  const now = options.referenceTime || new Date();
  const context = getProjectContext(options.projectRoot, now);
  const session = loadSession(context.sessionPath);

  if (!session) {
    if (options.dryRun) {
      return { sessionPath: context.sessionPath, payload: null, dryRun: true };
    }
    throw new Error(`No session file found at ${context.sessionPath}`);
  }

  const payload = {
    ...session,
    last_activity: now.toISOString(),
    last_command: options.lastCommand || 'heartbeat',
  };

  if (options.dryRun) {
    return { sessionPath: context.sessionPath, payload, dryRun: true };
  }

  await useLock(path.join(context.locksDir, 'session.lock'), async () => {
    await writeSession(context.sessionPath, payload);
    await appendOpsLog(context, {
      op: 'session_heartbeat',
      path: context.sessionPath,
      last_command: payload.last_command,
    });
  });

  return { sessionPath: context.sessionPath, payload };
}

function parseArgs(argv) {
  const args = { _: [] };
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg.startsWith('--')) {
      const key = arg.slice(2);
      const next = argv[i + 1];
      if (next && !next.startsWith('--')) {
        args[key] = next;
        i++;
      } else {
        args[key] = true;
      }
    } else {
      args._.push(arg);
    }
  }
  return args;
}

function asList(value) {
  if (!value) return [];
  if (Array.isArray(value)) return value;
  return String(value)
    .split(',')
    .map((item) => item.trim())
    .filter(Boolean);
}

async function handleCli() {
  const [command, ...rest] = process.argv.slice(2);
  const args = parseArgs(rest);
  const projectRoot = args['project-root'] ? path.resolve(args['project-root']) : process.cwd();

  switch (command) {
    case 'context': {
      const context = getProjectContext(projectRoot);
      console.log(JSON.stringify(context, null, 2));
      break;
    }
    case 'validate': {
      if (!args.schema || !args.path) {
        console.error('Usage: node shared/codex-helpers.js validate --schema <schemaPath> --path <jsonPath> [--warn-only]');
        process.exit(1);
      }
      const warnOnly = args['warn-only'] === true;
      const result = await validateJsonWithOptions(args.schema, args.path, { warnOnly });
      console.log(JSON.stringify(result, null, 2));
      if (warnOnly && result && result.valid === false) {
        process.exit(0);
      }
      break;
    }
    case 'save': {
      const tasksInProgress = asList(args['in-progress']);
      const tasksCompleted = asList(args.completed);
      const alerts = asList(args.alert);
      const testsRan = args['tests-ran'] === 'true' || args['tests-ran'] === true;
      const blockers = asList(args.blocker);
      const dryRun = args['dry-run'] === true;

      const result = await saveCheckpoint({
        projectRoot,
        status: args.status || 'open',
        sessionStatus: args['session-status'] || 'ACTIVE',
        testsRan,
        testsSummary: args['tests-summary'],
        testsDetails: args['tests-details'],
        testsRecommended: args['tests-recommended'],
        tasksInProgress,
        tasksCompleted,
        notes: args.notes,
        createdBy: args['created-by'] || '/save (fallback)',
        alerts,
        blockers,
        keepLatest: args['keep-latest'] ? Number(args['keep-latest']) : undefined,
        includeDiffstat: args['diffstat'] !== 'false' && args['diffstat'] !== false,
        dryRun,
      });

      if (dryRun) {
        console.log(`DRY RUN: would write checkpoint to ${result.checkpointPath}`);
      } else {
        console.log(`Checkpoint written to ${result.checkpointPath}`);
        if (result.retentionDeleted && result.retentionDeleted.length) {
          console.log(`Retention removed older checkpoints: ${result.retentionDeleted.join(', ')}`);
        }
      }
      break;
    }
    case 'end': {
      const dryRun = args['dry-run'] === true;
      const { sessionPath } = await endSession({
        projectRoot,
        sessionStatus: args['session-status'] || 'CLOSED',
        dryRun,
      });
      console.log(`${dryRun ? 'DRY RUN: would update' : 'Session updated at'} ${sessionPath}`);
      break;
    }
    case 'heartbeat': {
      const dryRun = args['dry-run'] === true;
      const lastCommand = args['last-command'] || 'heartbeat';
      const { sessionPath } = await heartbeatSession({
        projectRoot,
        lastCommand,
        dryRun,
      });
      console.log(`${dryRun ? 'DRY RUN: would update' : 'Heartbeat updated at'} ${sessionPath}`);
      break;
    }
  default: {
      console.log(`Usage:
  node shared/codex-helpers.js context
  node shared/codex-helpers.js validate --schema <schemaPath> --path <jsonPath> [--warn-only]
  node shared/codex-helpers.js save [--status open|closed|paused] [--session-status ACTIVE|PAUSED|CLOSED] [--tests-ran true|false] [--tests-summary "..."] [--tests-details "..."] [--tests-recommended "..."] [--in-progress "task a,task b"] [--completed "task c"] [--notes "..."] [--created-by "/save"] [--blocker "item1,item2"] [--keep-latest 3] [--diffstat false] [--dry-run]
  node shared/codex-helpers.js end [--session-status CLOSED] [--dry-run]
  node shared/codex-helpers.js heartbeat [--last-command "..."] [--dry-run]
`);
    }
  }
}

if (import.meta.url === `file://${process.argv[1]}` || process.argv[1] === __filename) {
  handleCli().catch((error) => {
    console.error(error.message || error);
    process.exit(1);
  });
}

export {
  getProjectContext,
  writeJson,
  writeMarkdown,
  validateJson,
  useLock,
  formatOutput,
  getOpenSpecStats,
  getActivityStats,
  getCoachingTip,
  saveCheckpoint,
  endSession,
  heartbeatSession,
  enforceCheckpointRetention,
  appendOpsLog,
  collectGitInfo,
  validateJsonWithOptions,
  loadSession,
};
