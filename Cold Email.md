# Cold Email Templates

## AI Generation Workflow

**Gemini API Key:** `AIzaSyCGy_RgCsDXxxXSgGvGqtz3d4qBipNUTNE`
**Model:** `gemini-2.5-flash` (free tier)
**Endpoint:** `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={API_KEY}`

### When the user asks to generate cold emails for a list of leads:

1. Read this file to load the API key and guidelines.
2. Read the Excel file provided by the user. Each row is a lead with a `Pain Point` column already filled in.
3. For each lead, call Gemini API with the prompt below — using the business name and the `Pain Point` value from that row.
4. Write `Subject`, `Email v1` and `Email v2` back into the same Excel file (or a new output file), keeping all existing columns.
5. Report the output file path when done.

### Gemini prompt template (per lead):

```
You are an expert cold email copywriter for a social media marketing agency called Y-Studio.
Write 2 different cold EMAIL versions for a restaurant/cafe called "{BUSINESS_NAME}".

The observed pain point is: {PAIN_POINT}

Use the PEC formula for both versions:
- P (Pain/Fear): Open with "I noticed..." referencing the pain point, then highlight the fear or desire behind it — what they're losing or missing out on.
- E (Evidence): Mention that you designed a demo specifically for {BUSINESS_NAME} and what result it could drive (more walk-ins, more sales, more leads, or stronger engagement).
- C (CTA): End with a soft, one-line ask — invite them to see the demo.

Rules:
- Format: cold EMAIL body only (no subject line, no sign-off)
- Salutation: start with "Hi {BUSINESS_NAME},"
- Tone: casual, direct, human — like a real person reaching out, not a pitch deck
- Length: 60–90 words each
- No hashtags, no emojis, no salesy buzzwords
- Never use words like "free", "guaranteed", "limited offer", "marketing agency", "social media services"
- The 2 versions must feel noticeably different — vary the fear angle, the evidence framing, or the CTA wording
- Do not use the word "just" or "I wanted to"

Return ONLY the 2 email bodies separated by "---", no labels, no numbering, no extra text.
```

---

**Formula:** PEC (Prospect's Pain/Fear → Evidence/Demo → CTA)
**Tone:** Casual, direct, human — not salesy
**CTA:** Always invite them to see the demo

---

## Subject Lines

Use the first available option in order. The subject must feel personal — like it was written for them, not blasted to a list.

| Priority | Subject | When to use |
|----------|---------|-------------|
| 1 | `Hi [First Name]` | Business owner first name is known |
| 2 | `Thoughts, [First Name]?` | First name known, want slightly more curiosity |
| 3 | `[Business Name] x Y-Studio` | No owner name available |

**Rules for subjects:**
- Never use words like "marketing", "social media", "services", "proposal", "collaboration" — triggers spam filters and sounds mass-blast
- Keep it under 40 characters
- First name always beats business name — do the research

---

## Templates by Pain Point

Pick the template that matches the lead's **primary pain point** observed from their IG/FB.

---

### Template A — Feed 全是手机随拍，画面杂乱

> **v1**
> Hi [Name],
>
> I noticed [Business Name]'s IG feed is mostly phone shots — and on Instagram, that's often what stops people from walking in before they even try the food.
>
> Most cafes that fix their visuals see a real jump in walk-in traffic and DM inquiries. I designed a demo for [Business Name] to show what that could look like — and what it could do for your sales.
>
> Mind if I send it over?

> **v2**
> Hi [Name],
>
> I noticed [Business Name]'s feed doesn't quite match the quality of what you're serving — and that gap quietly costs you customers every day.
>
> Cafes that level up their visuals tend to see stronger engagement and more walk-ins. I put together a demo specifically for [Business Name] to show what stronger content could do for your sales.
>
> Want to take a look?

---

### Template B — 超过一个月没有更新

> **v1**
> Hi [Name],
>
> I noticed [Business Name] hasn't posted in over a month — and when the algorithm stops seeing you, new customers stop finding you too.
>
> The cafes that stay consistent on IG tend to see a steady stream of new walk-ins and inquiries. I designed a demo for [Business Name] to show what getting back online could do for your sales.
>
> Mind if I send it over?

> **v2**
> Hi [Name],
>
> I noticed [Business Name] has gone quiet on Instagram — which usually means the algorithm has stopped pushing you to potential customers nearby.
>
> Staying visible is one of the easiest ways to drive consistent walk-ins. I put together a demo for [Business Name] to show what that could look like for you.
>
> Want to see?

---

### Template C — 更新频率极低（每月一次或更少）

> **v1**
> Hi [Name],
>
> I noticed [Business Name] posts pretty rarely — and with the algorithm rewarding consistency, low frequency usually means low reach, even when the food is great.
>
> Cafes that post consistently tend to see real growth in walk-ins and new customers. I designed a demo for [Business Name] to show what a stronger posting rhythm could do for your sales.
>
> Mind if I send it over?

> **v2**
> Hi [Name],
>
> I noticed [Business Name] only posts a few times a month — and most potential customers scroll past cafes they've never seen before.
>
> The ones that show up consistently are the ones people remember when deciding where to go. I put together a demo for [Business Name] to show what that could mean for your foot traffic.
>
> Want to take a look?

---

### Template D — Feed 全是 Reels，没有食物海报（粉丝 < 20k）

> **v1**
> Hi [Name],
>
> I noticed [Business Name]'s feed is all Reels with no food visuals — great for reach, but Reels alone rarely convert viewers into customers who actually walk in.
>
> Cafes that pair Reels with strong food content tend to see a real lift in walk-ins and orders. I designed a demo for [Business Name] to show what that could look like.
>
> Mind if I send it over?

> **v2**
> Hi [Name],
>
> I noticed [Business Name] doesn't have food visuals on the feed — and that's usually what makes someone decide whether to visit or scroll past.
>
> Strong food content is what turns reach into real customers. I put together a demo for [Business Name] to show what it could do for your walk-ins and sales.
>
> Want to see?

---

### Template E — 有海报但设计简陋

> **v1**
> Hi [Name],
>
> I noticed [Business Name] is already posting graphics — but the designs aren't quite matching the quality of the food you're serving. On IG, that gap is often what stops people from clicking through or walking in.
>
> Cafes that tighten up their visual branding tend to see stronger engagement and more sales. I redesigned a demo for [Business Name] to show what that could look like.
>
> Mind if I send it over?

> **v2**
> Hi [Name],
>
> I noticed [Business Name]'s IG graphics could be working a lot harder for you — design is usually the difference between someone pausing to look and scrolling straight past.
>
> I put together a demo for [Business Name] to show what stronger visuals could do for your foot traffic and sales.
>
> Want to take a look?

---

### Template F — 很少发食物内容

> **v1**
> Hi [Name],
>
> I noticed [Business Name] rarely posts food on IG — and most people decide whether to visit a cafe based on what the food looks like online.
>
> Cafes that show their food consistently tend to see more walk-ins and DM inquiries. I designed a demo for [Business Name] to show what that could do for your sales.
>
> Mind if I send it over?

> **v2**
> Hi [Name],
>
> I noticed food barely shows up on [Business Name]'s feed — which means potential customers scrolling through have no reason to choose you over somewhere they've already seen the food at.
>
> I put together a demo specifically for [Business Name] to show how food content could bring in more customers and increase your sales.
>
> Want to see?

---

## Excel Output

| Column | Content |
|--------|---------|
| `Subject` | Based on subject line priority table above |
| `Email v1` | PEC formula — angle 1 |
| `Email v2` | PEC formula — angle 2 |

---

## Placeholders Reference

| Placeholder | Fill with |
|-------------|-----------|
| `[Name]` | Owner first name, or business name if unknown |
| `[Business Name]` | Official cafe / restaurant name |

---

## Notes

- Pick **one pain point** per email — don't stack.
- If two pain points are equally obvious, default to **Template A** (visual quality) — most tangible.
- Avoid words like "free", "guaranteed", "limited offer", "marketing", "social media services" — triggers spam filters.
- The subject line is the open rate. Get the first name whenever possible.
- PEC order matters: lead with the pain they already feel, not your pitch.
