# Eitri - Offline Pandoc Document Converter

### Introduction

This is a side project that I've been working on during my time as a tech writer. Here's the context:

- Generally, most companies will employ single-source strategty to avoid duplicating technical content. These source files are generally written in Markdown.
- However, this content may need to be distributed in multiple forms (e.g. HTML, PDF, etc).
- At the same time, they may contain IP (Intellectual Property) or other proprietary/sensitive content which cannot be uploaded onto the cloud.
- Eitri is an **offline** tool which converts markdown and docx files into HTML, LaTeX, or PDF files. It's written in **python**, and utilises **Pandoc**.

### Main Features of Eitri

Eitri has a number of useful features:

1. Ablity to include disclaimer page.
2. Ablity to dynamically change the style of the title page.
3. Ability to customise a header to indicate the level of 'confidentiality'.
4. Ability to simultaneously generate an internal version and an external version of the same doc using <green></green> tags.
5. Ability to use custom latex syntax to further expand the features available in Markdown.

Eitri is built from the ground-up, and therefore offers maximal compatibility and flexibility. 

The code is built in a modular way - modifying parameters is extremely easy, and there is much scope for further enhancements.
