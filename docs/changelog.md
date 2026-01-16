# API Changelog

All notable changes to this API will be documented in this file.

## [Unreleased]

### Added
- New endpoint for user preferences management

### Changed
- Improved error messages for authentication failures

2026-01-15

### Added
- New `/api/v2/products/search` endpoint with advanced filtering options
- Support for pagination on all list endpoints
- Rate limiting headers in all API responses

### Changed
- Updated `/api/v2/users` endpoint to return additional profile fields
- Improved response times for `/api/v2/orders` by 40%

### Deprecated
- `/api/v1/products/list` endpoint - use `/api/v2/products` instead (will be removed in v3.0)

### Fixed
- Fixed bug where special characters in product names caused 500 errors
- Corrected timezone handling in date fields

## [2.0.0] - 2025-12-01

### Breaking Changes
- Removed deprecated `/api/v1/legacy` endpoints
- Changed authentication from API keys to OAuth 2.0
- Updated all date/time fields to ISO 8601 format

### Added
- New API versioning system (v2)
- Webhook support for order status changes
- Batch operations for product updates

### Changed
- Increased maximum page size from 100 to 500 items
- Renamed `customer_id` field to `user_id` across all endpoints

## [1.5.2] - 2025-11-10

### Fixed
- Security patch for authentication bypass vulnerability
- Fixed memory leak in file upload endpoint

### Security
- Updated dependencies to address CVE-2025-XXXXX

## [1.5.1] - 2025-10-22

### Fixed
- Fixed incorrect calculation in order totals endpoint
- Resolved CORS issues for subdomain requests

## [1.5.0] - 2025-10-01

### Added
- New `/api/v1/analytics` endpoint for usage statistics
- Support for gzip compression on responses

### Changed
- Increased rate limit from 1000 to 5000 requests per hour

---

## Legend

- **Added** - New features or endpoints
- **Changed** - Changes to existing functionality
- **Deprecated** - Features that will be removed in future versions
- **Removed** - Features that have been removed
- **Fixed** - Bug fixes
- **Security** - Security updates or vulnerability patches
- **Breaking Changes** - Changes that require client code updates
