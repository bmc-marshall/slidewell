# Launch checklist

Everything in the repo is ready. These steps need your GitHub account. 

## 1. Create the repo and push

```bash
git init && git add . && git commit -m "Slidewell 1.0.0"
gh repo create slidewell --public --source . --push
```

## 2. Description, website and topics

```bash
gh repo edit --description "Consulting-grade PowerPoint decks with Claude: agree the storyline headline by headline, then build a native, editable .pptx in your brand or template and check every slide before delivery. For recommendations, options appraisals, board papers, proposals and workshops. Includes a 63-check deck checker and 10 example decks."

gh repo edit \
  --add-topic claude-skill --add-topic claude-skills --add-topic agent-skills --add-topic claude-code \
  --add-topic claude --add-topic anthropic --add-topic powerpoint --add-topic pptx --add-topic python-pptx \
  --add-topic presentation --add-topic slides --add-topic slide-deck --add-topic pitch-deck \
  --add-topic presentation-design --add-topic consulting --add-topic strategy-consulting \
  --add-topic storytelling --add-topic pyramid-principle --add-topic workshop --add-topic ai-presentation
```

Then, in the repo's **Settings → General → Social preview**, upload `docs/images/hero.png`. This is the image shown when the link is shared.

## 3. First release

```bash
git tag v1.0.0 && git push --tags
```

The workflow builds `slidewell.zip` and attaches it to the release. That's the file claude.ai users download.

## 4. Check the installs work

- claude.ai: download the release zip and upload it in Customize → Skills.
- Claude Code: `/plugin install slidewell --marketplace bmc-marshall/slidewell`
- Other agents: `npx skills add bmc-marshall/slidewell`
- Checker: `pip install git+https://github.com/bmc-marshall/slidewell` then `slidewell-check some-deck.pptx`

## 5. Get listed

Each directory works differently and these change often, so check the site for its current process:

- **skills.sh** lists skills that people install with `npx skills add`; installs from step 4 start that.
- **SkillsMP** and **claudemarketplaces.com** index public GitHub repos that contain a `SKILL.md`; check that Slidewell appears, and use the site's submit link if it doesn't after a few days.
- **"Awesome Claude skills" lists on GitHub**: open a pull request adding one line, e.g. `[Slidewell](https://github.com/bmc-marshall/slidewell) - Consulting-grade PowerPoint decks: storyline first, every slide checked.`

## 6. Show it

The side-by-side image (`docs/images/side-by-side.png`) is the strongest thing to share: a short post with that image and the repo link, on LinkedIn and in communities where consultants and Claude users gather.
