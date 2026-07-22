'use strict';

// Pin the runner's timezone so the one locale-relative helper (formatRelativeDate)
// is deterministic wherever the suite runs. The London-anchored helpers set their
// own timeZone explicitly and do not depend on this.
process.env.TZ = 'UTC';

// Standalone Jest setup for @wbd/hub-core's pure logic (timezone/calendar
// derivations, cache and formatting). Deliberately decoupled from the SPFx
// `gulp test` pipeline: ts-jest transpiles the TypeScript directly (isolated
// modules, so type-only SPFx imports such as MSGraphClientV3 are elided and no
// SPFx build toolchain is needed to run the tests).
module.exports = {
  testEnvironment: 'jsdom', // provides window.sessionStorage for the cache tests
  roots: ['<rootDir>/src'],
  testMatch: ['**/__tests__/**/*.test.ts'],
  clearMocks: true,
  transform: {
    '^.+\\.ts$': [
      'ts-jest',
      {
        isolatedModules: true,
        tsconfig: {
          module: 'CommonJS',
          target: 'ES2019',
          lib: ['ES2020', 'DOM'],
          esModuleInterop: true,
          skipLibCheck: true,
          types: []
        }
      }
    ]
  }
};
