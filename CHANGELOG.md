<!-- do not remove -->

## 0.1.1

### New Features

- Move template marker rendering into converter as literal «body» runs; tmpl callable now takes token node dict ([#8](https://github.com/AnswerDotAI/mdhtml2docx/issues/8))
- Move `jinja_literal` to mdhtml.jinja and import `mustache_kind` from mdhtml.mustache ([#7](https://github.com/AnswerDotAI/mdhtml2docx/issues/7))
- Add details-div degradation to bold label and switch dash handling to explicit replacements callback ([#6](https://github.com/AnswerDotAI/mdhtml2docx/issues/6))


## 0.1.0

### New Features

- Migrate from JustHTML to fast5ever, using isinstance node checks and direct attrs access ([#5](https://github.com/AnswerDotAI/mdhtml2docx/issues/5))
- Rename to mdhtml2docx and switch input from xhtmlmd XHTML to mdhtml JustHTML DOM with support for raw-data encodings, reference tokens/groups, HTML5 repair, inert templates, and figcaption alt text ([#2](https://github.com/AnswerDotAI/mdhtml2docx/issues/2))
- Add xhtml2docx conversion engine, WML helpers, style management, validation, Word scripting, and syntax highlighting modules ([#1](https://github.com/AnswerDotAI/mdhtml2docx/issues/1))
- Support MDHTML raw-data encodings, reference tokens and groups, HTML5 body repair, inert templates, and figure accessibility text from `figcaption`.
