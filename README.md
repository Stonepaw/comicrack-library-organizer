Library Organizer is a rule and token based file organizer plugin for ComicRack.

It uses a templates to insert various Comic Book fields to rename and/or move your comics. It supports multiple sets of templates and can be run on specific type of books based on rules.

## Library Organizer 3.0

Library Organizer 3.0 is in progress. This involves a rewrite into c# to make it much easier to maintain and improve. I have no deadline on this except that I am actively working on it as of Oct 3, 2026. This means that extensive changes to the python files or forms may be rejected. Please open a feature request to later add such features to 3.0.

This is planned to have initially feature parity with bug fixes and improvements (HDPI!). Currently the only major changes planned will be add most of the remaining comic book fields into the rules and template engine.

## Contributing

Contributions are welcome as this is open source software.

New features and bugs should have a ticket created first for discussion. I will try to answer promptly but feel free to `@` me (repeatedly if necessary) if I do not. Open bugs and features can be claimed, just note on the ticket that you are working on it.

However we need to talk about LLM usage.

### LLM usage

My stance on LLM usage follows from the [Rust LLM guidelines](https://forge.rust-lang.org/policies/llm-usage.html) which is a very great read. LLMs are a tool, not a developer.

Important parts:

- It’s fine to use LLMs to answer questions, analyze, distill, refine, check, suggest, review. But not to **create**.
- Comments and PR descriptions must be written by the submitter. I want to hear from you, in your own words, what you have created. If I want an LLM summary of the PR I can easily generate one myself.
- I reserve the right to reject obvious LLM generated PRs without comment. This is a passion project for me. I won't spend time reviewing code that you haven't.
- LLM usage must be disclosed.
