SAMPLE_DSL = """\
// Document setup
page size=letter orientation=portrait margin=72,72,72,72

// Title
h1 align=center color=a4 | "Word DSL Demo Document"
h2 align=center color=a2 | "Generated with the Word Copilot"

// Horizontal rule
hr color=a3 weight=2

// Intro paragraph
p spacing_after=6 | "This document was built from a simple DSL. You can mix " + "bold" bold=true color=a1 + ", " + "italic" italic=true + ", and " + "colored" color=a3 bold=true + " text inline."

// Bullet list
ul indent=1
  item | "First bullet point with " + "emphasis" bold=true color=a2
  item | "Second point — plain text"
  item | "Third point with " + "italic detail" italic=true

// Numbered list
ol indent=1
  item | "Step one: open the app"
  item | "Step two: write your DSL"
  item | "Step three: click a button"

// Divider
hr color=a3 weight=1

// Table
table width=100% header_fill=a4 header_text_color=#FFFFFF text_color=t1 border_color=a3 border_weight=1 header_bold=true
cols=30%,40%,30%
header="Feature","Description","Status"
row="Headings","h1–h6 with color + align","✓ Done"
row="Paragraphs","Rich inline text","✓ Done"
row="Lists","ul / ol with nesting","✓ Done"
row="Tables","Header + row styling","✓ Done"
row="Images","Remote URL embed","✓ Done"
row="Icons","Font Awesome 6","✓ Done"

// Spacer
br

// Icon + caption
icon name=circle-check style=solid width=40 height=40 color=a2 align=center
p align=center color=a3 | "Icons are downloaded from Font Awesome and embedded as images."

// Page break to second page
---

h1 color=a1 | "Page Two — SVG & Floating Textbox"

p | "Below is an inline SVG triangle, followed by a floating textbox anchored to the page."

// Inline SVG
svg width=80 height=80 color=a2 align=center
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
  <path d="M12 2 L22 20 L2 20 Z" fill="currentColor"/>
</svg>
endsvg

// Floating textbox
textbox left=480 top=120 width=240 height=80 color=a6 outline=a3,1.5 padding=10,10,8,8 halign=center valign=middle | "Floating sidebar note" size=11 bold=true color=a4

p spacing_before=8 | "The textbox above is anchored to the page at an absolute position, independent of the text flow."
"""
