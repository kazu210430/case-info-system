# Small Policy Inventory

This inventory tracks the small policy / small test units that have been extracted from larger services. Each policy is intended to keep pure logic out of service orchestration and make fragile branches easier to lock down with focused tests.

## TaskPaneManager Policies

TaskPaneManager now exposes the reusable decision points around pane host reuse and post-action refresh behavior as small policies.
| Policy | Protected branch | Test |
| --- | --- | --- |
| `TaskPaneHostReusePolicy` | decide whether an existing case host can be reused | `TaskPaneHostReusePolicyTests.cs` |
| `TaskPanePostActionRefreshPolicy` | decide whether to refresh, skip, or defer after an action | `TaskPanePostActionRefreshPolicyTests.cs` |

### Nearby Unprotected Area

- pane visibility orchestration and host lifecycle sequencing remain in `TaskPaneManager`

## DocumentCommandService Policies

DocumentCommandService now separates the precondition, route, and execution-decision branches into small policies.
| Policy | Protected branch | Test |
| --- | --- | --- |
| `DocumentCommandPreconditionPolicy` | decide whether command execution should continue or be blocked | `DocumentCommandPreconditionPolicyTests.cs` |
| `DocumentCommandActionRoutePolicy` | decide whether an action routes to document, accounting, caselist, or unsupported handling | `DocumentCommandActionRoutePolicyTests.cs` |
| `DocumentCommandExecutionDecisionPolicy` | decide whether a precondition result leads to continue or throw | `DocumentCommandExecutionDecisionPolicyTests.cs` |

### Nearby Unprotected Area

- message construction and service-to-service execution sequencing remain in `DocumentCommandService`

## KernelWorkbookService Policies

KernelWorkbookService now isolates the small decision points around startup -> promotion -> window restore -> home release fallback, while leaving workbook state collection and UI orchestration in the service.
| Policy | Protected branch | Test |
| --- | --- | --- |
| `KernelWorkbookStartupDisplayPolicy` | startup time: decide whether HOME should be shown | `KernelWorkbookStartupDisplayPolicyTests.cs` |
| `KernelWorkbookPromotionPolicy` | home release: decide whether the kernel workbook should be promoted | `KernelWorkbookPromotionPolicyTests.cs` |
| `KernelWorkbookWindowRestorePolicy` | home release: decide whether global Excel window restore should be avoided | `KernelWorkbookWindowRestorePolicyTests.cs` |
| `KernelWorkbookHomeReleaseFallbackPolicy` | home release: decide the fallback after promotion is not taken | `KernelWorkbookHomeReleaseFallbackPolicyTests.cs` |

### Nearby Unprotected Area

- workbook enumeration, ActiveWorkbook resolution, and actual Show/Restore orchestration remain in `KernelWorkbookService`

## Next Candidate Checklist

- keep the extracted unit as pure logic and leave orchestration in the service
- avoid COM / VSTO / UI side effects in the policy itself
- prefer a single `bool` or a very small enum when that is enough to express the branch

## Next Candidate Shortlist

1. `TaskPaneManager`: a remaining pane visibility branch that can be separated without pulling in UI side effects
2. `DocumentCommandService`: a small unsupported-action decision that can be isolated without moving message formatting
3. `KernelWorkbookService`: follow-up fallback branches near the startup / promotion / window restore flow, if a single minimal decision remains unprotected
