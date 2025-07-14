const { parseCompanyFromDomain, parseCompanyFromSenderName } = require('../ParsingUtils.js');

describe('ParsingUtils', () => {
  describe('parseCompanyFromDomain', () => {
    test('should return null for invalid sender', () => {
      expect(parseCompanyFromDomain('invalid-sender')).toBeNull();
    });

    test('should return null for ignored domains', () => {
      expect(parseCompanyFromDomain('<test@.io>')).toBeNull();
    });

    test('should parse company from domain', () => {
      expect(parseCompanyFromDomain('<test@google.com>')).toBe('Google');
    });
  });

  describe('parseCompanyFromSenderName', () => {
    test('should return null for invalid sender', () => {
      expect(parseCompanyFromSenderName('<test@google.com>')).toBeNull();
    });

    test('should parse company from sender name', () => {
      expect(parseCompanyFromSenderName('"Google" <test@google.com>')).toBe('Google');
    });
  });
});
