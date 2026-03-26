/**
 * Pure helper: validates that a string looks like a GitHub PAT.
 * @param {string} token
 * @returns {boolean}
 */
const _validatePATFormat_test = (token) => {
  if (!token || typeof token !== 'string') return false;
  const trimmed = token.trim();
  return /^ghp_[A-Za-z0-9]{36,}$/.test(trimmed) || /^github_pat_[A-Za-z0-9_]{20,}$/.test(trimmed);
};

/**
 * Pure helper: builds Gist URL from an ID.
 * @param {string} gistId
 * @returns {string}
 */
const _buildGistUrl_test = (gistId) => `https://gist.github.com/${gistId}`;

function test_validatePATFormat() {
  // Valid classic token
  Assert.isTrue(_validatePATFormat_test('ghp_ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijkl'), 'PAT: valid classic ghp_ token');
  // Valid fine-grained token
  Assert.isTrue(_validatePATFormat_test('github_pat_ABCDEFGHIJKLMNOPQRSTU'), 'PAT: valid fine-grained github_pat_ token');
  // Invalid: empty
  Assert.equal(_validatePATFormat_test(''), false, 'PAT: empty string rejected');
  Assert.equal(_validatePATFormat_test(null), false, 'PAT: null rejected');
  Assert.equal(_validatePATFormat_test(undefined), false, 'PAT: undefined rejected');
  // Invalid: wrong prefix
  Assert.equal(_validatePATFormat_test('gho_ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890'), false, 'PAT: wrong prefix rejected');
  // Invalid: too short
  Assert.equal(_validatePATFormat_test('ghp_short'), false, 'PAT: too-short classic token rejected');
  // Whitespace trimming
  Assert.isTrue(_validatePATFormat_test('  ghp_ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijkl  '), 'PAT: whitespace-padded token accepted');
}

function test_buildGistUrl() {
  Assert.equal(_buildGistUrl_test('abc123def456'), 'https://gist.github.com/abc123def456', 'GistUrl: builds correct URL');
  Assert.equal(_buildGistUrl_test(''), 'https://gist.github.com/', 'GistUrl: empty ID returns base URL');
}
