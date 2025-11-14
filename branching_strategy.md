# Git Branching Strategy for Customer-Specific Releases and Hotfixes

## Scenario
- Monthly releases (v1.0, v2.0, ... v10.0, v11.0, etc.)
- Customer wants to stay on Month 10 release
- Customer needs a hotfix on Month 10 without upgrading to Month 11+

## Branching Strategy

### Visual Representation

```
main (development)
│
├─────v1.0────v2.0────v3.0────...────v10.0────v11.0────v12.0───► (continues)
│      │       │       │              │        │        │
│      │       │       │              │        │        │
│      │       │       │              │        │        │
│      │       │       │              │        │        │
│      │       │       │              └────────┘        │
│      │       │       │              Customer          │
│      │       │       │              stays here        │
│      │       │       │                   │            │
│      │       │       │                   │            │
│      │       │       │                   ▼            │
│      │       │       │              release/v10.0     │
│      │       │       │                   │            │
│      │       │       │                   │            │
│      │       │       │              [Fix needed]      │
│      │       │       │                   │            │
│      │       │       │                   ▼            │
│      │       │       │           hotfix/customer-v10.0-fix1
│      │       │       │                   │            │
│      │       │       │                   │            │
│      │       │       │                   ▼            │
│      │       │       │              [Testing]         │
│      │       │       │                   │            │
│      │       │       │                   │            │
│      │       │       │                   ▼            │
│      │       │       │              v10.0.1 (Tag)     │
│      │       │       │              (Customer         │
│      │       │       │               Hotfix)          │
│      │       │       │                                │
│      │       │       │              [If needed later] │
│      │       │       │                   ▼            │
│      │       │       │           hotfix/customer-v10.0-fix2
│      │       │       │                   │            │
│      │       │       │                   ▼            │
│      │       │       │              v10.0.2 (Tag)     │
│      │       │       │                                │
```

### Detailed Flowchart

```
┌──────────────────────────────────────────────────────────────────┐
│                        MAIN BRANCH                                │
│  (Ongoing Development - Monthly Releases)                         │
│                                                                    │
│  Month 1 ──► Month 2 ──► ... ──► Month 10 ──► Month 11 ──► ...  │
│     v1.0      v2.0              v10.0         v11.0              │
└────────────────────────────────┬─────────────────────────────────┘
                                 │
                                 │ Customer freezes at v10.0
                                 │
                                 ▼
                    ┌────────────────────────────┐
                    │   release/v10.0 branch     │
                    │   (Long-term support)      │
                    │   Created from v10.0 tag   │
                    └────────────┬───────────────┘
                                 │
                                 │ Bug found - needs fix
                                 │
                                 ▼
                    ┌────────────────────────────┐
                    │ hotfix/customer-v10.0-fix1 │
                    │   (Branched from           │
                    │    release/v10.0)          │
                    └────────────┬───────────────┘
                                 │
                                 │ 1. Make fix
                                 │ 2. Test
                                 │ 3. Review
                                 │
                                 ▼
                    ┌────────────────────────────┐
                    │  Merge back to             │
                    │  release/v10.0             │
                    └────────────┬───────────────┘
                                 │
                                 ▼
                    ┌────────────────────────────┐
                    │  Tag: v10.0.1              │
                    │  Deploy to customer        │
                    └────────────────────────────┘
```

## Implementation Steps

### Step 1: Create Monthly Releases (Normal Flow)
```bash
# On main branch
git checkout main
git tag -a v10.0 -m "Month 10 Release"
git push origin v10.0

# Continue with next month
git tag -a v11.0 -m "Month 11 Release"
git push origin v11.0
```

### Step 2: Create Long-Term Support Branch for Customer
```bash
# When customer wants to freeze on v10.0
git checkout v10.0
git checkout -b release/v10.0
git push origin release/v10.0

# Protect this branch in GitLab
# Settings → Repository → Protected Branches
```

### Step 3: Create Hotfix for Customer
```bash
# Create hotfix branch from release/v10.0
git checkout release/v10.0
git checkout -b hotfix/customer-v10.0-fix1

# Make your fix
# ... edit files ...

git add .
git commit -m "Fix: Description of the bug fix"
git push origin hotfix/customer-v10.0-fix1
```

### Step 4: Create Merge Request in GitLab
```
Source: hotfix/customer-v10.0-fix1
Target: release/v10.0

Title: "Hotfix: [Description]"
Labels: hotfix, customer-specific, v10.0

Reviewers: Assign appropriate reviewers
```

### Step 5: Merge and Tag
```bash
# After MR approval and merge
git checkout release/v10.0
git pull origin release/v10.0

# Create patch version tag
git tag -a v10.0.1 -m "Hotfix 1 for v10.0 - [Description]"
git push origin v10.0.1

# Deploy v10.0.1 to customer
```

### Step 6 (Optional): Cherry-pick to Main
```bash
# If the fix is relevant to current releases
git checkout main
git cherry-pick <commit-hash>
git push origin main
```

## Branch Naming Convention

| Branch Type | Naming Pattern | Example |
|-------------|----------------|---------|
| Main Development | `main` or `develop` | `main` |
| Release Branches | `release/v{major}.{minor}` | `release/v10.0` |
| Hotfix Branches | `hotfix/{customer}-v{version}-{description}` | `hotfix/customer-v10.0-fix1` |
| Monthly Tags | `v{major}.{minor}` | `v10.0`, `v11.0` |
| Hotfix Tags | `v{major}.{minor}.{patch}` | `v10.0.1`, `v10.0.2` |

## GitLab Configuration

### 1. Protected Branches Settings
```
Settings → Repository → Protected Branches

Branch: main
- Allowed to merge: Maintainers
- Allowed to push: No one

Branch: release/*
- Allowed to merge: Maintainers
- Allowed to push: No one
```

### 2. Tag Protection
```
Settings → Repository → Protected Tags

Tag: v*
- Allowed to create: Maintainers
```

### 3. Merge Request Approvals
```
Settings → Merge Requests → Approval Rules

For hotfix/* branches:
- Require 1-2 approvals
- Require pipeline success
```

## CI/CD Pipeline Example

```yaml
# .gitlab-ci.yml

stages:
  - test
  - build
  - deploy

# Run tests on all branches
test:
  stage: test
  script:
    - pytest tests/
  only:
    - branches

# Build for monthly releases
build_monthly:
  stage: build
  script:
    - ./build.sh
  only:
    - tags
    - /^v\d+\.\d+$/  # Matches v10.0, v11.0, etc.

# Build for hotfixes
build_hotfix:
  stage: build
  script:
    - ./build.sh
  only:
    - /^v\d+\.\d+\.\d+$/  # Matches v10.0.1, v10.0.2, etc.

# Deploy to customer
deploy_customer:
  stage: deploy
  script:
    - ./deploy_to_customer.sh
  only:
    - /^v\d+\.\d+\.\d+$/  # Only hotfix versions
  when: manual
```

## Benefits of This Strategy

✅ **Customer Stability**: Customer stays on v10.0 without forced upgrades
✅ **Isolated Fixes**: Hotfixes don't include unrelated changes from newer releases
✅ **Traceability**: Clear history of what was deployed to each customer
✅ **Flexibility**: Can support multiple customers on different versions
✅ **Minimal Risk**: Changes are isolated and tested independently

## Multiple Customers on Different Versions

```
main ──► v10.0 ──► v11.0 ──► v12.0 ──► v13.0
          │         │         │
          │         │         └── release/v12.0 (Customer C)
          │         │                │
          │         │                └── v12.0.1 (hotfix)
          │         │
          │         └── release/v11.0 (Customer B)
          │                │
          │                └── v11.0.1 (hotfix)
          │
          └── release/v10.0 (Customer A)
                 │
                 ├── v10.0.1 (hotfix 1)
                 └── v10.0.2 (hotfix 2)
```

## Best Practices

1. **Document Everything**: Maintain a CHANGELOG.md for each release branch
2. **Customer Tracking**: Use labels/tags in GitLab to track which customer uses which version
3. **Security Updates**: Even for frozen versions, apply critical security patches
4. **Deprecation Policy**: Set a timeline for how long you'll support old versions
5. **Communication**: Keep customers informed about their version status
6. **Automated Testing**: Ensure comprehensive tests for each hotfix
7. **Release Notes**: Maintain separate release notes for hotfix versions

## GitLab Issues/Milestones Structure

```
Milestones:
├── v10.0 (Month 10 Release)
├── v10.0 - Hotfixes
│   ├── Issue: Bug XYZ fix for Customer A
│   └── Issue: Security patch for Customer A
├── v11.0 (Month 11 Release)
├── v12.0 (Month 12 Release)
└── ...
```

## Summary

This branching strategy allows you to:
1. ✅ Continue monthly releases on `main` branch
2. ✅ Create `release/v10.0` branch for customer who wants to stay on Month 10
3. ✅ Create `hotfix/*` branches for specific fixes
4. ✅ Tag hotfix releases as `v10.0.1`, `v10.0.2`, etc.
5. ✅ Optionally cherry-pick fixes to main if relevant
6. ✅ Scale to multiple customers on different versions

This is a proven strategy used by many enterprise software companies for managing customer-specific releases and hotfixes.
