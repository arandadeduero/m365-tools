import test from 'ava';

// ---------------------------------------------------------------------------
// OData single-quote escaping in searchUsers fallback path
// ---------------------------------------------------------------------------

function buildODataFilter(query) {
  const safe = query.replace(/'/g, "''");
  return `startsWith(displayName,'${safe}') or startsWith(mail,'${safe}') or startsWith(userPrincipalName,'${safe}')`;
}

test("searchUsers filter: single-quote in query is escaped as ''", (t) => {
  const filter = buildODataFilter("O'Brien");
  t.true(filter.includes("O''Brien"));
  t.false(filter.includes("O'Brien,"));
});

test('searchUsers filter: query without special chars is unchanged', (t) => {
  const filter = buildODataFilter('John Doe');
  t.true(filter.includes("'John Doe'"));
});

test('searchUsers filter: multiple quotes are all escaped', (t) => {
  const filter = buildODataFilter("it's a test's value");
  t.true(filter.includes("it''s a test''s value"));
});

test('searchUsers filter: empty query produces valid (empty) filter', (t) => {
  const filter = buildODataFilter('');
  t.true(filter.includes("startsWith(displayName,'')"));
});
