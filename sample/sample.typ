#import "@preview/rubby:0.10.2": *
#let ruby=get-ruby()
#let counter=counter("counter")
#set text(
  font:("Libertinus Serif","Noto Serif CJK SC"),
  size:22.5pt
)
#set page(margin:(top: 0.2cm,rest:0cm))
#table(
  columns:(auto,auto,auto),
  stroke:0pt,
  inset:10pt,
  align:horizon
,[#counter.step()#context counter.display()],[#ruby("なん","何")#ruby("ど","度")も#ruby("じ","事")#ruby("こ","故")に#ruby("あ","遭")うなんて、#ruby("かの","彼")#ruby("じょ","女")には#strong[#ruby("しんそこ","心底")]#ruby("どうじょう","同情")する],[\u{2460}]
)
