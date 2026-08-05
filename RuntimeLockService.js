/**
 * Serialization for the explicitly protected runtime workflows.
 * Wave R1.3.0.2 — Runtime Safety (Reduced Scope)
 */

const FO_RUNTIME_LOCK_TIMEOUT_MS_ = 5000;
let FO_RUNTIME_LOCK_DEPTH_ = 0;
let FO_RUNTIME_CONTEXT_ = null;
let FO_RUNTIME_CONTEXT_SEQUENCE_ = 0;

/**
 * Returns the immutable Runtime Context for the current execution.
 * The context is created once and reused by all downstream consumers.
 */
function foRuntimeContextGet_() {
  if (FO_RUNTIME_CONTEXT_) {
    return FO_RUNTIME_CONTEXT_;
  }

  FO_RUNTIME_CONTEXT_SEQUENCE_ += 1;

  const startedAt = new Date();
  const platformVersion =
    typeof FO_CONFIG !== 'undefined' &&
    FO_CONFIG &&
    FO_CONFIG.PLATFORM_VERSION
      ? String(FO_CONFIG.PLATFORM_VERSION)
      : 'UNAVAILABLE';

  FO_RUNTIME_CONTEXT_ = Object.freeze({
    runtimeId:
      'RUNTIME-' +
      startedAt.getTime() +
      '-' +
      FO_RUNTIME_CONTEXT_SEQUENCE_,
    executionMode: 'PRODUCTION_RUNTIME',
    authorityLevel: 'FULL',
    platformVersion: platformVersion,
    startedAt: startedAt.toISOString()
  });

  return FO_RUNTIME_CONTEXT_;
}

/**
 * Test-only reset hook.
 * Production workflows must not use this function.
 */
function foRuntimeContextReset_() {
  FO_RUNTIME_CONTEXT_ = null;
}

function foAssertRuntimeLockHeld_(operationName) {
  const operation = String(operationName || '').trim();

  foAssertRuntimeSafety_(operation);

  if (FO_RUNTIME_LOCK_DEPTH_ < 1) {
    throw new Error(
      'Runtime safety blocked unlocked operation: ' + operation
    );
  }
}

function foWithRuntimeLock_(operationName, callback) {
  const operation = String(operationName || '').trim();

  foAssertRuntimeSafety_(operation);

  if (typeof callback !== 'function') {
    throw new Error('Runtime lock callback must be executable');
  }

  if (FO_RUNTIME_LOCK_DEPTH_ > 0) {
    FO_RUNTIME_LOCK_DEPTH_ += 1;
    try {
      return callback();
    } finally {
      FO_RUNTIME_LOCK_DEPTH_ -= 1;
    }
  }

  const lock = LockService.getScriptLock();
  let acquired = false;

  try {
    acquired = lock.tryLock(FO_RUNTIME_LOCK_TIMEOUT_MS_);
    if (!acquired) {
      throw new Error(
        'Runtime safety blocked concurrent operation: ' + operation
      );
    }

    FO_RUNTIME_LOCK_DEPTH_ = 1;
    return callback();
  } finally {
    FO_RUNTIME_LOCK_DEPTH_ = 0;
    if (acquired) lock.releaseLock();
  }
}
