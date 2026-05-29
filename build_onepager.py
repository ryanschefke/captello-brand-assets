#!/usr/bin/env python3
"""
Captello Onboarding One-Pager builder (HTML). Larger fonts + progress visuals.
Usage:
    python3 build_onepager.py content.json out.html
Inlines logos/captello-horizontal-blacktext.svg from CWD. No external libraries needed.

content.json schema:
{
  "client": "Philips",
  "call_type": "returning",          # controls progress NOW stage + recap/goals label
  "owner": "Amanda",
  "status_line": "Returning client · 2 working sessions remaining · owner: Amanda",
  "goal": "One-sentence engagement goal.",
  "challenges": "One sentence of top challenges.",
  "package": ["ULC x15","QuickScan","Salesforce integration"],
  "session1": {"title":"Set up the foundation","items":["...","...","..."]},
  "session2": {"title":"Prepare for the floor","items":["...","...","..."]},
  "contacts": [["Caroline Salter","Primary — email addressee"], ...],
  "next_steps": ["Confirm ...","Send ...","..."],
  "tickets": [["#38416","Waiting on customer"], ...]   # optional; [] to omit
}
"""
import json, sys, os, html

def esc(s): return html.escape(str(s))

def build(c, out):
    svg = ""
    p = "logos/captello-horizontal-blacktext.svg"
    if os.path.exists(p):
        svg = open(p).read()
        if "?>" in svg: svg = svg.split("?>",1)[1].strip()

    first = c.get("call_type") == "first"
    now = 0 if first else 1
    stages = ["Kicked off","Working sessions","Event-ready"]

    # progress tracker (horizontal)
    seg = ""
    for i,st in enumerate(stages):
        cur = "cur" if i==now else ("done" if i<now else "")
        seg += f'<div class="stage {cur}"><div class="dot"></div><div class="slabel">{esc(st)}</div></div>'
        if i < len(stages)-1:
            seg += f'<div class="bar {"fill" if i<now else ""}"></div>'

    pkg = "".join(f'<span class="chip">{esc(s)}</span>' for s in c.get("package",[])) or '<span class="chip">See contract</span>'

    def sess(card):
        items = "".join(f"<li>{esc(i)}</li>" for i in card.get("items",[]))
        return f'<div class="card"><div class="lbl">{esc(card.get("label",""))}</div><div class="ttl">{esc(card.get("title",""))}</div><ul>{items}</ul></div>'
    s1 = dict(c.get("session1",{}), label="SESSION 1")
    s2 = dict(c.get("session2",{}), label="SESSION 2")

    contacts = "".join(f'<div class="c"><div class="n">{esc(n)}</div><div class="r">{esc(r)}</div></div>' for n,r in c.get("contacts",[]))
    steps = "".join(f"<li>{esc(x)}</li>" for x in c.get("next_steps",[]))

    tickets_block = ""
    if c.get("tickets"):
        chips = "".join(f'<div class="tk"><span class="tkid">{esc(tid)}</span><span class="tkst">{esc(stat)}</span></div>' for tid,stat in c["tickets"])
        tickets_block = f'<div class="sec"><h2>Open Tickets</h2><div class="tks">{chips}</div></div>'

    recap_label = "Goals" if first else "Where We Left Off"

    doc = f'''<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{esc(c.get("client",""))} — Onboarding At a Glance</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700;900&display=swap');
*{{margin:0;padding:0;box-sizing:border-box;}}
body{{font-family:'Roboto',Arial,sans-serif;color:#000;background:#fff;}}
.page{{width:8.5in;min-height:11in;margin:0 auto;padding:0.55in 0.65in;}}
.top{{display:flex;align-items:center;justify-content:space-between;border-bottom:4px solid #FF0000;padding-bottom:16px;}}
.logo svg{{height:38px;width:auto;display:block;}}
.eyebrow{{font-size:13px;font-weight:700;letter-spacing:2.5px;color:#FF0000;text-transform:uppercase;}}
h1{{font-size:42px;font-weight:900;margin-top:22px;letter-spacing:-0.5px;line-height:1.05;}}
h1 .accent{{color:#FF0000;}}
.status{{font-size:17px;color:#333;margin-top:10px;font-weight:500;}}
.status b{{color:#000;}}
/* progress tracker */
.track{{display:flex;align-items:center;margin:26px 0 8px;}}
.stage{{display:flex;flex-direction:column;align-items:center;flex:0 0 auto;width:130px;}}
.stage .dot{{width:22px;height:22px;border-radius:50%;background:#fff;border:3px solid #D9D9D9;}}
.stage.done .dot{{background:#000;border-color:#000;}}
.stage.cur .dot{{background:#FF0000;border-color:#FF0000;box-shadow:0 0 0 5px rgba(255,0,0,0.15);}}
.stage .slabel{{font-size:14px;margin-top:8px;font-weight:500;color:#555;text-align:center;}}
.stage.cur .slabel{{color:#FF0000;font-weight:700;}}
.stage.done .slabel{{color:#000;}}
.bar{{flex:1;height:4px;background:#D9D9D9;margin-bottom:24px;}}
.bar.fill{{background:#000;}}
/* sections */
.sec{{margin-top:26px;}}
.sec h2{{font-size:16px;font-weight:700;text-transform:uppercase;letter-spacing:1.2px;border-bottom:2px solid #FF0000;padding-bottom:6px;margin-bottom:12px;}}
.gc{{font-size:16px;line-height:1.55;color:#262626;}}
.gc b{{color:#000;}}
.chips{{display:flex;flex-wrap:wrap;gap:10px;}}
.chip{{background:#000;color:#fff;font-size:15px;font-weight:700;padding:8px 16px;border-radius:22px;}}
.row{{display:flex;gap:20px;}}
.col{{flex:1;}}
.card{{background:#F7F7F7;border:1px solid #EAEAEA;border-top:5px solid #FF0000;padding:14px 16px;height:100%;}}
.card .lbl{{font-size:12px;font-weight:700;color:#FF0000;letter-spacing:1.5px;}}
.card .ttl{{font-size:19px;font-weight:700;margin:4px 0 10px;}}
ul{{list-style:none;}}
.card li{{font-size:15px;color:#262626;padding-left:18px;position:relative;margin-bottom:8px;line-height:1.35;}}
.card li::before{{content:"•";color:#FF0000;font-weight:700;position:absolute;left:0;}}
.contacts{{font-size:15px;}}
.contacts .c{{display:flex;padding:6px 0;border-bottom:1px solid #EEE;}}
.contacts .c .n{{font-weight:700;width:42%;}}
.contacts .c .r{{color:#444;width:58%;}}
ol{{counter-reset:s;list-style:none;}}
ol li{{counter-increment:s;padding-left:30px;position:relative;font-size:15px;color:#262626;margin-bottom:9px;line-height:1.35;}}
ol li::before{{content:counter(s);background:#FF0000;color:#fff;font-size:12px;font-weight:700;width:20px;height:20px;border-radius:50%;display:flex;align-items:center;justify-content:center;position:absolute;left:0;top:1px;}}
.tks{{display:flex;flex-wrap:wrap;gap:10px;}}
.tk{{display:flex;align-items:center;gap:8px;background:#FCE9E9;border:1px solid #F3C6C6;border-radius:6px;padding:7px 12px;}}
.tk .tkid{{font-weight:700;font-size:15px;color:#000;}}
.tk .tkst{{font-size:13px;color:#B30000;font-weight:500;}}
.foot{{margin-top:32px;border-top:1px solid #DDD;padding-top:12px;font-size:10px;color:#888;text-align:center;letter-spacing:0.3px;}}
</style></head>
<body><div class="page">
  <div class="top">
    <div class="logo">{svg}</div>
    <div class="eyebrow">Onboarding · At a Glance</div>
  </div>
  <h1>{esc(c.get("client",""))} <span class="accent">·</span> Working Session Prep</h1>
  <div class="status">{esc(c.get("status_line",""))}</div>

  <div class="track">{seg}</div>

  <div class="sec"><h2>Your Package</h2><div class="chips">{pkg}</div></div>

  <div class="sec"><h2>{esc(recap_label)} &amp; Challenges</h2>
    <div class="gc"><b>{"Goal" if first else "Recap"}:</b> {esc(c.get("goal",""))}<br><b>Top challenges:</b> {esc(c.get("challenges",""))}</div>
  </div>

  <div class="sec"><h2>The Two-Session Plan</h2>
    <div class="row"><div class="col">{sess(s1)}</div><div class="col">{sess(s2)}</div></div>
  </div>

  <div class="row" style="margin-top:26px;">
    <div class="col sec" style="margin-top:0;"><h2>Key Contacts</h2><div class="contacts">{contacts}</div></div>
    <div class="col sec" style="margin-top:0;"><h2>Immediate Next Steps</h2><ol>{steps}</ol></div>
  </div>

  {tickets_block}

  <div class="foot">CAPTELLO | 13101 PRESTON RD. STE 110 – 159 DALLAS, TX 75240 | 888.399.6430 | CAPTELLO.COM &nbsp;·&nbsp; Internal onboarding reference</div>
</div></body></html>'''

    open(out,"w").write(doc)
    print("WROTE",out)

if __name__=="__main__":
    c=json.load(open(sys.argv[1])) if len(sys.argv)>1 else {"client":"Sample","call_type":"returning"}
    build(c, sys.argv[2] if len(sys.argv)>2 else "onepager.html")
