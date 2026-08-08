/**
 * Tests for credential storage. Pure — no PropertiesService, no Sheets.
 */

function test_isRealSecret() {
  Assert.isTrue(_isRealSecret("ghp_abc123", "PASTE_GITHUB_TOKEN_HERE"), 'realSecret: an actual token counts');
  Assert.isTrue(!_isRealSecret("PASTE_GITHUB_TOKEN_HERE", "PASTE_GITHUB_TOKEN_HERE"), 'realSecret: the placeholder does not');
  Assert.isTrue(!_isRealSecret(SECRET_MOVED_NOTICE, "PASTE_KEY_HERE"), 'realSecret: the moved-notice is not re-migrated');
  Assert.isTrue(!_isRealSecret("", "X"), 'realSecret: empty is not a secret');
  Assert.isTrue(!_isRealSecret("   ", "X"), 'realSecret: whitespace is not a secret');
  Assert.isTrue(!_isRealSecret(null, "X"), 'realSecret: null is safe');
}

function test_planSecretMigration() {
  const specs = [
    { name: "a", cell: "B9",  label: "RapidAPI Key", placeholder: "PASTE_KEY_HERE" },
    { name: "b", cell: "B13", label: "GitHub PAT",   placeholder: "PASTE_GITHUB_TOKEN_HERE" },
    { name: "c", cell: "B20", label: "Unused",       placeholder: "NONE" }
  ];

  const plan = _planSecretMigration(specs,
    { a: "rapid_live_key", b: "ghp_token", c: "NONE" },
    { a: null,             b: null,        c: null });

  Assert.equal(plan.move.length, 2, 'migrate: both configured secrets are moved');
  Assert.equal(plan.notSet.length, 1, 'migrate: an unconfigured secret is reported, not moved');
  Assert.equal(plan.conflict.length, 0, 'migrate: no conflicts in the clean case');
}

function test_planSecretMigrationIsIdempotent() {
  const specs = [{ name: "a", cell: "B9", label: "RapidAPI Key", placeholder: "PASTE_KEY_HERE" }];

  const second = _planSecretMigration(specs, { a: SECRET_MOVED_NOTICE }, { a: "rapid_live_key" });
  Assert.equal(second.move.length, 0, 'idempotent: a second run moves nothing');
  Assert.equal(second.alreadySecure.length, 1, 'idempotent: reported as already secure');
}

function test_planSecretMigrationRefusesToGuessOnConflict() {
  const specs = [{ name: "a", cell: "B9", label: "RapidAPI Key", placeholder: "PASTE_KEY_HERE" }];

  const clash = _planSecretMigration(specs, { a: "cell_value" }, { a: "stored_value" });
  Assert.equal(clash.conflict.length, 1, 'conflict: differing cell and stored values are flagged');
  Assert.equal(clash.move.length, 0, 'conflict: nothing is overwritten');
  Assert.equal(clash.alreadySecure.length, 0, 'conflict: not silently treated as done');

  const same = _planSecretMigration(specs, { a: "  same_value  " }, { a: "same_value" });
  Assert.equal(same.conflict.length, 0, 'conflict: identical values differing only in whitespace are not a conflict');
  Assert.equal(same.alreadySecure.length, 1, 'conflict: matching values count as already secure');
}

function test_secretSpecLookup() {
  Assert.isTrue(_secretSpec("githubPat") !== null, 'spec: the GitHub PAT is declared');
  Assert.equal(_secretSpec("githubPat").cell, "B13", 'spec: legacy cell retained for migration');
  Assert.isTrue(_secretSpec("rapidApiKey") !== null, 'spec: the RapidAPI key is declared');
  Assert.equal(_secretSpec("nope"), null, 'spec: an unknown name returns null rather than throwing');
}
