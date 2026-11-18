# Spec Delta

## REMOVED Requirements

### Requirement: React and Vue adapter parity

**Reason**: This fork ships neither the React adapter nor, after this change, the Vue adapter, and the `check:parity-contract` script no longer exists.

**Migration**: The requirement "Header, footer, and notes chrome" in `angular-adapter-chrome` covers the adapter side. Platform-agnostic header and footer logic stays in `packages/core`.
