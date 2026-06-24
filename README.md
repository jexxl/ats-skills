# ATS Skills

This repository packages reusable Codex skills for Autonomous Transportation Systems (ATS) work.

## Included Skills

- `algorithm-packaging`: package Python algorithm models as Docker deliverables for cloud platforms or data submission.
- `ats-algorithm-cover`: generate ATS algorithm cover/card HTML and transparent high-DPI PNG images.

## Use

Clone the repository:

```bash
git clone https://github.com/jexxl/ats-skills.git
```

Then copy or link the skill directories into your Codex skills directory so they remain direct child directories:

```text
skills/
├── algorithm-packaging/
└── ats-algorithm-cover/
```

If your skills directory is also a Git repository, you can keep this repository as a submodule and expose each skill with links or junctions from the skills root.

## Update Workflow

Commit changes in this repository first, then update any parent repository that tracks it as a submodule.
