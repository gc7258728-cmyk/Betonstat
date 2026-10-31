# -*- coding: utf-8 -*-
"""BétonStat : statistiques exploratoires, données et méthodes explicites."""
from __future__ import annotations

import base64
import hashlib
import io
import json
import math
import re
import unicodedata
import zlib
import zipfile
from datetime import datetime, timezone
from html import escape

import altair as alt
import numpy as np
import pandas as pd
from scipy import stats
import streamlit as st
from openpyxl import Workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pathlib import Path

APP_NAME = 'BétonStat'
VERSION = '2.0'
ACADEMIC_TITLE = 'Conception et développement d’une application web d’analyse statistique de la résistance à la compression du béton à partir de données expérimentales.'
SOURCE_URL = 'https://archive.ics.uci.edu/dataset/165/concrete+compressive+strength'
DATA_URL = 'https://archive.ics.uci.edu/static/public/165/data.csv'
DATA_SHA256 = '8d4b15b6fc68cd932d745cbd663d5ceae66dd54422e99c1e4865f2936ab7e2af'
SOURCE_CITATION = 'Yeh, I. (1998). Concrete Compressive Strength [Dataset]. UCI Machine Learning Repository. DOI : 10.24432/C5PK67. Licence CC BY 4.0.'
RETRIEVED = '2026-10-31'
# Copie sans modification du CSV officiel, compressée pour une utilisation hors ligne.
# Attribution : SOURCE_CITATION. Le décodage doit reproduire exactement DATA_SHA256.
EMBEDDED_UCI = 'c-oa+%dRB1k=+~cKj;Grh2)!DdJOmnyqL8hG-Ls`EVSeSX1;#noD<Bdy0<cu*M?+wGm1$D<HdOKuiyUl?bknU|MkxwfB)l$zy0U?uRs3u?T5er^T(fWfBWbE{_vOI|8e{8AOHCF{r30&{PykpZ-Dycr(gch{^GCy{_*?o-+uVZ&p&_v_OpT9{`SkS&wu>O&)*E@*Pp(B`{UaWKmGe(zkUDq`|rQ}&$l0b|Kt0&Uw{7NAAkN+$=eP84K?24z1+fPe^Fb##eQq!KE^wcQYJ_Z_ii95aUkR0+b!C_b{i?)%E-6awp-rrv7LjZT5dVqxANydh0*Wz77RG!vm*kmb^B9~#eSNq&6w}$&wq;5;6i?xtlIa60kr#O5UJar4!)M|?GV8Jkm@~`<8Rm>a=CYd$gL1TpJy2NIIj0&t4;mBwQES5Ga6*HZ^mYS&g}@{>AnTbhgr6__=)!f{A~GVNBM4Z?cRg%^LBy~?*pH(O&Td|<7vr8MD5nMdLP`|xTk6p<!IadG!ssu)O*N>M=?l0r5+(Vbnd2}V}(#gt@(1vLUOzYOgB4=dI`oa=X*RvCQ)EK5sv_Hq9xsr$(ctU?x7r_8H=!QgbUI!nr-)-j!@k|sP&L)xRG*?^_XjXucVv{V)qkk@sPtB4?d)0Lg(kwy&4zBgdOYO?`E={<!}LJ=-KXDKbK@6`F65wF<(sQ-@@)q0>$Ga1rKibTu)E*A?VIN2B<WS$d9y>?&0)9gNIe}DUi`m5&rU!PWjAJp->4NP&$cXySIA2p>~Uvba*h0EhOH1JU>Vt!ML}6OkZ4%t^49IK@y3~cfdKo-nfTxDNLHJ^<EZ~3IS_AB0^WUBQ2=g+aQk}ES=P`-AB1F13FN!9ixCk&a`?zQMmvyiGam~?TLU!3%72Is9bDs#|OqK4Ddum@k~2ygZGTbIm(LYcT45M^l6YP@r25tjkkLTr*|D@$UtBSakDQGpamLhsex|I?!a){?!lI0{~E>);nqyi+Q)7w-=S`UsTAXr=)gE@Qo;s`Hq*|ur2(5R>|kmbe${?ppe79LPZjXcDYxF6{X5}+^GDEd$AWCyk*cjVX#thW%qVlDJ9ZjueTK&bqT`ffM4l$>Cs1&(ISAMTK4|L`T(dueM2+pR%4+I1G;Iif1vZa!m9QcV6vpV6Hk3NokG}$JgH9)uR%%SrgeQ}bB?V52BLe<lNz6%yHgA8QzV?0YJqFQEDEOW+ciMqX3t$0-``Yl?VMNC@G@u`6Fzisf|AK3PIZ}aSBNtlt8T$&@DM;3crSSYsV&kkL79V@`Q^2tqWf{8}Nw-@qgTmXk%Yz-*kDYz=-3^F!OrNmsx)s6JFRdYVpet@BY`8Z-1bicGo5{MJx}wN?ed3&Kun}P$GJ#0NFWT0)Ets&NgF$mnU&nr8{4k%<*4OrwQhUtV01DK-LfU@95@$E)qPWP}4OYk+XE!LPxWL(5KwshPhRai1;q3-lTbFpdL4k}fZE1zK8wBQ9;B9XIi{D@2?KUOr0&f!sYrIX&jz!+)g8Le8H&}mPTf-V}<NoobX)idt@#?w2*#<c7+gSW|j6~}tYttNFV{Q3bxPletHt;eQxf@e&bW~jOw|jh67#u68YaEVQsW%o_9I&fA_So18mt(;j7JSL(+FJhV8df+RQ`>>AvAXb9mWWao;;JGCw9M88ZpWG<W!V>&vBvM%VdFZ!F!cq`lY$mlUO4zFb}xAzw*0uSS?Z<UYFakc0lljJ1e+FEexLQ4<pC<&`jX`hw8cf1$FyV7WZw6+!tn-NXft1OJRn!EL60S@Fg&2Cf>szFSWAoS0NA$7@CC<B9$sL0U_&JFg=MTUyy#+laqbI_2V9V^|AOZMITkMu4+LNGJP<Y%D7j|(5hlZ}Fg>6v&Espz1Z-Kd<27qrVS0eZds$(612y!($+f`s0H!+r%Enf>-hgeOE^s{{$D*&`v4j<_2XtTMdSId~Dk8v+)hjsmSHC*(8Xp5(H^DDVeZj_nOH21~gEVnhzlgC*HpZ#Z%@m`2r33s5cD9Q@U@Iu0Hwf2+FO@Y|^FhL8QI>!aoywFY3&19L;#i&84v?;O20_I011A~=5R-|wMw*P+-X>S=6(lwn8Q!}D#nvJtbeEvmSd$m`asjq2BYDIR!(_B=BM%oiA{fy&=7yKPC;}0QvL$Gp-587D*kWO9XbIf5Y1<pCOXpTiCuO>l1#h3r2;D;pQ}YkFjI#$?Rv@v(AaC~-D7MF&(Iq5yR@PSPbXgF?qBI9=AYYTlLC_mSy6NK~*0Td#g~skomc>PIvENvP#xC=Hb#5M?#S_SY)JvvI?5_tjKgoLoH=8u-^o<(}ZfOZto80;mu(p`jfU!e`{Z^KN4L(ji`KcR^npXkx79?yc+$7Q#J(wd-TmY7!5o;}gE6ShSA~^QhU5KnKMl_3<YJh!D{R49@v`Wk{jRrVreWOP{D5$vvjja}*&=NSdsxw*whitV!rPal33~p-r&VW3B&`Ioll6i+3^}X7azN|vqUd>QMsCErr)Ek#qgM*Ds<5^sfn3UoBKvM^&9*GvSa_vGzTB!8z^}@(Z=*z&i-6XnM{DIPtsI=F9fXj%{T3cVz#_+!{slWLFP*f|y5!f-Px#LS{Ki<3uu1M%a{w28W6`tr-aB{8jGH_#JtGPUjlxG%-<?$oVXY#B`1SjS;EkR2_rML(#LZ)P00hb_7WpyzUP@NW3CzqRr5ny@K&pv<$yNV&67KSG{oWe;cDbwZ@P+2tbdPZz%Fxlbi)7zd9SfGRkxwhjr#r(Pgn(j5oa7~mZmymARt$^aFOeCy;N{^(KrE*tMM>?;Q8*Y!Gz*^{}-3_MFusfGiBk5%Y8r!p*Tv!Chp#|oXEP;dZ3~Y7jzSU`k)fae3Ac>*mKj_YRX%$?wnc6D6=fSOlW4C!7oVzt0s2|Xj$oT6r-0-E?ZDUB94rT^kLPB$dNiFA~*d7B1myp<*GPbV2uU(P1QgD|MtzGkZNzfu{VzzGAzU-%Y309;^jdcyIJfoJDfJLTj6LWt>s)|0d-PrA%iuiy@NeN52*<2<pKx2<CW>>%^acWrvht#RkYfFpSm_bbU#T)oaf_~kqvfGxR-QAHE!6Ti6yy~xjYpXS{E@s=4b72w?zS>{c!3%a}wmWZVZHgTv8ZW@=+FX}lah#15z5uMNb6Ep+s!fEB_`jxD%v}cOvw1tIW+<j$-8S0hDXzh4gJlg^`Xeqb1ALiQr{)&zKGJBy!~a9KPlTH@6rTMiVTJ!`<^hY`XZ`{__G$|XU-*f39O4&#B7-NCuPjWP$8X};ri%~vQdmqt(x5TXD^srMslNn-D#=9XB`9{8rOK{ApLEIP0<<d%tR`>%pl`e)$5Vli2A6hkMQMUbF2)65+gw)F$_g|t9G%ITE6})G1$xz%R<bdVl(LF;Ak9IjpTJO$`4wy!(ThQnMOsN&g2e$Anb_CBDpVP3tH2?4dRbPMv9ai2b74$3EnOv<_mL>W$YKWew@h>kOb}jx#RmQAbqy?&V63e|wp>^Pev+;c`P+c*G?2QWapi;<?%h+{x#UYwh#Qk=eFaw6DL62$VP)^oX?VScjmg#<aUALL5^%0`o7(w4@6FDx^%YR?Ut~tiHLP-SmR6u(eosqFu&q_4tVagg<jsS#iH<H%Zf5|FVJ?an+TaVozP&|g<<2awL-`eHbs^ZK**gGlW%oYoT2kC)iP{)i3Pro@i@@N&E|Yu(9NP@KudRUNEUQR*X)&PGF2DKkpmN|q!)-;E6GWEG=1baXU<J|=G`3n9;aLJ#<y>Bbmcl@QXKgLmY4w|9l9#730$Z}uI-*cxIhd`=_t-pgvH~u#+R`e#+;KEMmf#_O+&Eub&&Fzv&Ao}me-QZtYE)j*!0)ZO^Rf9oe+At4y9zE}>&UC{I^CR>mxJvp%HBW?gLebEE9>@%v)4dxbX4KPrugpS4<MeJ+{KH(K)&3=x?1^#a+>G}SY4H56$LWAmO@aJ?}kQiAlEnMV~?mj|0e0S`U_b_WVY3xcs^$M!cVY!)#9u#4BIyXA)z4ipv(ZoU<QeWvuQ3ej|7(<I{Q=zy!r!4?7IGgJ*{)~2cY)k=v`cs#92+k)o=$>ph1FcB3^kKgD)e`>cEs&!R0vG$|}5jrsFC+?rV*g+Qs#3tQL&S&#9`)LzSL9h7@}>N@`c$IZmP6;T|u*bPpZv0!)6BtChKD2u5MqyaLbK1K{pvHgQRwv0K`#J1jT_<96<T*kERe&U`P#LMX`pQXbFE(~~yK5~J8bh^F8}n5d#OPLAXRZeh`zrSz~Ez69@3uw(r(OJm|@33YE`z&x*wp<0UcK#ftN?+4X=%mgd^$1}Jw2aEEMom3bl8}^G9H_r!$>G_zWXpV*yhH3w>_>l#C!7}ZPNGvE1m$2csPf5f5Zg=e`BG!m1KXaFvP)WM1;d+8>slxE(F&v7PKtrR^6Uh(4QxqYDEb+&XnEbPq#hWq~5dM@$iAh((#94`M){rI>Lw4E#(xt$X&+{P?gY|7?$Ay<F)-kii!SS0GK1&`0NTAluGo(@Y{Ca5V&7t8pAwZWRbt^Mm+ie+^X6HWAz@R!fv&)IpBfWq<+CObLqUeZ*fk}$Tbpm_u$YR?jE)}u~Dkz2v9OhZCdj%$P!`h?X7K@evm+eq0R}HxU>0(6So36=yb_pqgLXSMJsc2gOvQJWB+EGX+(R-m1c_0?wVw|Ca?&X#-0}4^_olT8LFo_|jI#myKL#RBT%s5Q-T+eNz!1n7;$ae=cEMiNEVRG8$T`_Tni(%U-VgbDz8ZgdlGSf>#V~wEXeuEO=Ee6SiG&zmtAYmfy`>0>Al35q+>c6O35UjA!D3ds){xA6OmbIkHq)S=BriIvvqtWPQ9KR`#6NYD<iB!=9b((jfjrkQw{-X;AE>?3nN{n;)%`-O&4;2%3B+KfaGPopUuz0L1K2~3p&+s9Uq*6uFG_df>=rb3MQM9LE;ZR-z61bIN+vpuKpZ54{XwnTMLo@5f6s8`n;Zs@}9juT6V<1?i0$`*reU1c};W9jOpovktt+dC{7pMw>e><$cOjokix3NrE@?t6zWRbEILqpq>Rnk)2hUqMewB`E<E5rp@){?CfmzGn?5_RDS#HyF%rI-zaFiQmH>liB(rn%!Pi7A&&0^@?ll#UpeNep_WDrswt#H1OdVirkEx+|_zn7o(u3(`yB-MB(vT1^h?6sFL44@(s0$;vANOd&AY-YXO)L!;&u3e%cnTcj{yoUy7a3X|2$w7@kI(<P~`t4BH9$w{juCbvWT0*xu8*uNmfd>?Iv$OMZ<HLFCXO)IPVlE{?#)K<ip0&F#}5}7a=>k5%6X35OG6(ZA8mWa&Ly2m<|d1Pr@C$pLQ_7ytw#K2m=pf0r)L^hW7rJMv?T~e3=?q3jS?qeKzO=o(xYh6*Ao4RtXYfZ(=#uB~F%!OsWsf<53YqNEaiOAd0#epR;Ky8SAH%El%90>VN@p2~Jc|QPceqEeEVw#hg)GMQ2z{K+IV;^7=6QoXH8UvQ$0w(pVZ#BO%o`A?vh_-*{K;9ki0;urDe1d9KZ8nxcHhLbw7+#z<tj$rEzr=VU=1)L~T0a4SbNVMBuIUQs`E!hEw{rS5o@@!HA7e|Si>)<i>VCd6b|o37pK4&~dj2ULS~~xha+60p{hDY5sxuY%3~q4dz~_;WGhrv&TX1${x64<c=*RCsMk5FnWUFI~bV=K_*rXZc6`UetJUOABfz0bdp!fi4%+)-ZiKm7e<_V12nv7EzmClxNx}r2MJtr`YDm0$L2;1zl6_KnBC3}Mcg~A0;EAsS2I+fo6bpa*Z+bBk!U}{%M>|-L+Na6Hjjcxf(0I*~=(mDPR7Zc7G)0kI3PGL~-oRz__K?0#})^KRJL>WPrQ&q(#%-lH!VwbjZ3es4gc0NsEJu>88>|h=OnSsMp>i$?z_L$}|5EJ|IDM(`<rt@i(Ji<nvaC{E|6(#|XXA~I)!Qi-nLXz7!ha#VacnOqf^_~H-84YFeIgkSFXF#CIemRw*gvc4mX+|7i`u4=^%rHF7HHlLCbHm6%A9o+LxHj@yJA<KHicap4o<Y#Ai?tu_wv1;GgE?vabQU=T*r@NNKVhOh8@$RjD;wz;h8?FEL^%W*>XsPKVD?sM>1Dl%X?8hIXJqBosXT#Lk_^kz*_${@BpyLD#_i)7`qhi_cnY}YonQ`E`V_`GglTL~aJCZF-<H!^)X9{V%i%dU=IiF;Ex{j)+gD*X0JvimIC7(2Q?U}iAuh_KdBjIk@m7)!I<ORD6hWfuFRC1d@f>fq@OIXqVjScMo6c_W11wY7T5x}d?M4?M`j5u;vRu1!fTXtR8;Oe8mf|;)^fEEqM6~FOXf!7Ec`fB&5x7Koj{g-NJY>VlC<fGX`+#0Ml$P^;-lVM2X7KJaw!Aqhja8Maya3|n+hEU22~gGqrb`DekesWu?gh0wV33#pngXOUIqN9*{^9Ebce_@{<RcXzT;g5_#n?CFw;?&u4%{dxichR8q*E-3{%Ec1zWD?f*!-dOxu2+L#QA=7I+PvEBaW}gWYdM(9!K65PMlV@`lR7(bWZP1NM2I-Ux*oD=@36nNg!T__PlZS`)G4YL{J$dRgahF`Rd?&c~S?0DhuD1aEvR=DSlfzah(zQm=kL9qiJ0QEXb0^TL-YJo)RR~V-f{(rfY6`IGwKug==LU+2yvJ>GLH$tLJiFlDi5yu)TS&p1mJlfO{hSYvOBCalLm*dsM?rjVYfbSm2#<XHFNi7oyH<;I3z+@IW#86~|F{V!{Q~h?VLR9dYX<qiZ!IbSB~5H8)pTfk?jy#CSB!><Jkm?}ibZY9gt(<1UmB#|Qf(gyZ?u6rM8p{qj-f(N*2xR<(H&9iVsEB~`sK-i{(7I6JVs8<#AS3G!}S)*{pR_VM&QKhiSTp7i}VmLI)Fq`4*<VFHxuHAj44;O)3Xir0<28RtbH9sV5n54upk|9t0VLy639=dat$C@n5gMxWy~@#4#x+qa`YU-d$TcjH7~bL;Ev^uV8SH`L$OJ|Y+y;+@TK@*w~TsOlw&_W(>u$VbDnB}+OQ*KKXYBMN{u%-qSDw$haw6pMS`yjj5eVdnVg{Nj{YLGmgwFNm@#w8L?U9eVC|fGm-Wq(xV;<#=2vghh$*4v-|=keJ`4Oyc3NSN#Kh+sph0-bRsKPYdJyxbzP%srv_%#g~vo-t{s9f++{%9z14mviSv<xMil)=|PEuadxv;{mdL!Ub}&JWX$1+S~?uZ4mM^|r5!*kgqx~ulEKVGdq1v{G!~f94oX|0nLMqmfu5SoAmp8+XsI6X>Cb=-sacgx2lS#4@}}XV746QeBWs|0(4`zPKp1vNP`jyU`Y|8-Zf7r3j%o?0uWAL)u!Y~&C7m5_LsT=fMD4Z%7Ns+8S|pyC;H{4h-=fOVAFxBP=v2-r;ae-|>HaF`8k>W)<8cV8qKNM0qP!u?(oVfntBAovdVG0QV7YH9-VC0zurc0M)!J;IkXBmH!%3-7*zxx75D!Wyq=cZW4Nv1@O6*+SUzm2-o>sWw?%zl7SC16?x}pP_ZR4D*JOzaBO{6_(l2FKjH_0{j42s@X$eEO%6z^vO<-^IL#)FmZsi7<xgvadt>8sdkqq#?Y&@O;HgOb-_L?Yu8MZsT^^eBNU$a<d-X*eR_&hqN(N<a41X-Y(?JzZr`<?&Dm{X{H&qa0D9;xP`$d)}x)XZeg!+N<=Q!&#lqzo-5j?V35IBTHK07=D@QX9uhbjag68Dv95(_DCW;{*E+f?C0ooVbHIUAV@w3cS7@n_Hel%KfRR1l(Cb!=s93(=nL{<{%OePr|k-(UyBs`6fMd#3DB4<InIu;?Ny4W#|VP%;R;I6wzV^U)q%6z;{*-5@i5KvL|)1pL3wmR7P{*#2x`!Mi`O=QGT_J|7m%>s%Y_Ukc2BuYv{=%)ildoHJM{!ewq{<kUzqrkR`j*qA}%Dgt|_cTJz1S>A=u()Cl7>O2@T1^u(C}(+c*}OSn-p3E;#5-VxPSe-UMptX+K0*L-j(tHU`9X)bl4OBMUGO`eHh0)V_4)vO>y=!qKtvf^)$ygZ1O^8@@i)1%C_+6==uts`7q(4<8(iyTOwV0pw0;$16bXgq*-Bg%<!^D5HydT-XoBxx}0g$K`Hn#{haPa2$VuOY};f50B=tX}c%$^S<<c99L?jA{-36ad`}&_cy5+-veNYk6v=mS2EPEn0r!pKE@*6kGtcXRnzt5PB&1`!U?*UY`-0s--l%idjzApZx)_e-Ac#1z^pkDr^N^b`+i)~BKGG9$;<pMco=*PFI&WGcP+aU`FXfJ4|~kysbI%d3|t#l(sg5hhI%}n=Dq%4oXIx0)*sEE=LQY{Jj>|$L0}1}#}~{`3_3a(w=b9)2R9x7sD>&X&oEFzMoK*TgdLs<K!j6|z5xe9mW#E_ebNDiGipVY#^(;-4y%b{IUJWvq66vexH@u=el)I|V=9N=mj|-#rG%5vP=MLsBZNY2N28Qyhodwh<MGt=ZjZxp_pO(YvTPt{Df-&S(J0?a43hV6UkCNNuj2`LzvAlRD~mRSG#Gw39Q0_%=U>}@5a;kHM?2*22Ltb(d2#p#@l~?N{OaMju-Daw=MzBnr$77Q@RJBk{YAutek5cazwo=(@^DSQA6A%}J?Gv7$T#x&gi&VsdO&Z_k9Qot+4^=^_{oRVB~8kz;`ik9;OzI~x~h<~-UG-yqwoAZfWKJG=W7p0xb9CDr!9Wq?;hBrV6O80c&G6Vi=ylv=<!hoR>eRMB)P7p@fd(lx~h-WkuuN|sNq-?^NmTh;ft@-xbss|ABy+mic{9{hF`WPetDbJMMpLhy67!>J1mz}=||(r_F^QZzbx=boRs(C<gvX2adNr+4=>$rc<ENWIu3_Dr6(SZ4~Bu5<Mei%qPppY__qKIbamqQfDZY3X+Iq1^OHc(5gBLT|Cr2`syvPX2Acx4;}w90@%#lyB!1#G+-y~8*p9{@68GTueVbK~@0nQ=^+R^jN2^lr4!$=_ncxS3TYLd}7r2kge;Qp4;b>SdoH5#l+xLbG2NSSi^V=OxHS|f22-Xweed64i_G*@UN61M191!!wwG~O88E@FwM5s&0bs7cd<hQ`bW&#>ETmCFGl3``2ggy=@c~&45B{~N@2vZSr#yc#teE`~2_@&W`rob`;!W%GlVzq?llF6LC1%bL1;6iThA@P@z=@Or~@QxQZV8&EYoKAN_v>9ua{|VHaW?KMyDB0y=s1pIdWw+}zdqYI8cS3iJG+dUl?#Qyv`+PKG-jlJVO}pN>A%9mR>$%v?dnwyuDk0E{R~;r0cB9`Wcvz5CQic^ganqu2nGR@l<iKEJMH14hJ7SxP@T07fH?7Xz92Qz+qP$_oj>LNh6j$W9pw{fTDeQOQ#t<r?_ylO30m4B?NMz?#+WJK|I`%M7@spKk-DE5D{b9lGLew~<RpuL*zE6+qb-F*A7!*Yl*$OS<0EVY5Z1k_9HYR!0IItl)^8ROiXqufOI@`vPSAiLuq=jvIGP5wlkSm3Hm3<S0JKfPz@E5blP8AKkjbxhl%VJBtj<X3Y`hcHE95-ftFq6YkrBJJ?#%?N!Xk+t7IBDppT}5V|6W59q?bhdm;Y@C`gU3p%XRF&<prU51ElGHVJX}ZFS_+=tn!V{5AgS5O7U6P)$}<|TS0^cD+o=F8v40BrzU%?<t_qf@PHtUi6htk%kYSZ&<~vOUiiR&>&R#lGV#n-g+1XWyXki=ptxxgcmeu2YfPE4-E`7U+4Xw*gF*Z`!Ri-LAU`%(g9<xA@bDPWvlCoXbK23)#cbW0-4C4ba1?1a&OC)a)ZJ_bU)*;%p8oqe>p0o#rhnTttBumT-74rlkT@rif@T?)G;zV9_SqCX3R^vB~@waLC!?5L(_bQPDiEM$g7YSF>P0x^qVHCb<V-@1xq?XHFe-y-?rQ1*1mO16vaM4UjS7us+Ep^B$s-w+RM(zit8`6kuUJ+5HG=(l1WYl?x5HfaD@D+)+aSw>@h!2v?d@KfT|1ob0o39m>P@At6v=0g7U20s_&?cb7?0D49zEnMvB)HsGo0)ezMmG5*+{DI1<P*GrhQiC_SYg&1BGU39<9rgIF0hy89Qkrub2rc%BshxP7W*uO$nTc(w5h1(Qqa+xgg1NsXvOX_MbM^5x&Q`BckHrbM3)`)Y3JKx-y4g8q^aRkj4Vam`KU;HrkPyXlmLBG?<QIVcw7(eb%Jzpnrp1$b-g%s6NWnLpvv*c+Xa|Ns~rg1tEmp24s`V5syiZyuSo^(9KseYFc<LjApRGR#HnpjPM(gzcu}Q)7%XB<Z7Ipp?%v3rISJg?;E17~Mq%<UGG`%CoYJGwZnh?1j@`lc=IWmQVzdj<E0fSDlgYNiJh~EfGQgmkiH;OvcB7PZB_8DqsPaPL`jKP+Q6eW627_E^Q|uWi?)wIeP6-fun<aE-3*u<Ewmx%oR3gX+rtc=4d!3Zah+X0&Rrh3It%kW0YzTHJ2H~4;56te3ypnpAK+`-JV2}Z9*DkxgyV`^AAy{50*9){Df&9w)@c^*i?5IO%6F2+DL(F0K^{6P<ISu*@hg1*-#G{4N*isKfH1)*mfii(Cv}J`1)qT#4q&y4?D8J;Sn_cV4H2lPiy9{!dgH9m(G09@sX2oVyI|Wq6fTT88djsZJ-BIT|@|Imt0uyz8W;Vy2GMfzeu2x#|Kv@{6<f$Xx$&JVd&xuxuj>vui^PXU-zfP80`%H9C6gBqR8&r1kK|^c?8Y0h#M%SH%#t$@0kZOg=m!=M<KuyxSLz1vkEW70r^=JZ51rAUOzDBtbW-GfWrxU-$rbd})3Cd=hm#B@N3WgB@gpO((tpTign_XUbjF4kRM<|Z9V51cb&bQ}^$_98}quGZJZ3t}cf=tbAVGcx?3sCPY-Rl`oU+*2GyOVe!EkGynxwhT-^Hz(z)A89_Nbt^iaDk4#o}>m#G(9*Gvp6B-tE|kZzzX$i<we1|A&~aczO6`+2n(Hj-=+2Ogd-Z-O{&OB9Ohe~XsXD^<MkLWAA5ut@Hep@#>)N4TtTTO$Pr;wv*i3;0ykRNf|e=hlws@@5tI{8)BL_0FuYjuIY;Dr%7{TIlx^i$ZtZC`(L1Hfl;I3&WQ?Enz^s$j_a=>Q)OZrTBFV^MXV5+K#e=<yt?Q1!{q2H*`!jn0f!+A5Z151Q`=9v8DH;W;y&udSf0^=O@UZn%ImOyEd=eXEBt3eU{ApM|uP{xZdww;7aA<A@u}KAf+7<4fJy+cvifpouRkMwI2;TkKtRi48&S$cA%jnIzU$F78k5qY(s`?o6(KodZ!O``!y(GApt)%J$JrN5FWMiLgJW-1luCwP7AI&T3HS7I58XpBTM5CV7q|g1;eByD49Jm0|FU=_IFg=nY-IQ|bWVce5$kL!{<c6qYoJQ573?ngTLM4GTuJ97d2lGk^H=R9NhaPBR9|EKpwCc-FO0Jfv)%f}|pRdA~`3f>myY=}{?TgvPZfV2~U;ONR@qxUR)1DEl#9%dDQMu)v1`HRDU$CeN3mPv58UBp!!eWHx-DZsf+5f~JJt}(Gm9|;%N78C1wZgx!k-_pkzFKK}?F;i&u!au0?ZFe;V6V}sl0~d3`4gXAZx|0g#oS4~A8a@8;?7Pr-WHgJah2`ApRG7_%T^Z{=d}TbCW#{%4R+K$z1VbmCw7SKNid{D*>-}3br+}Tqay+9I2L`=(t9r{JDWX&-`b1S=ZsMG(wTaEyY1&qcOTXt+u~q#U0u-C;D-;~;=EAe_r?|kfwRaL+~yXcEvT{t?Aa=O;Y`G7KXc}89&)+P7KdlQn~FTa#J2{R{Up9;#~%0-U^ed6VSM0ailELmPF?NoS^Q1|(o3L<GR0^@Ry{!t?x`{(!NzsIQsVD54WD=OR<3M5jb9cZ=*g&T@3qHAIsY&&@6mIJuv?YZ!4iYDKA0Xf^V*+HP4DuqcFDj7i@wa{^yYKb1qz!>&1V0NNzeC(@p&(==BFr+hi#_KomC-jhWfI1Lp0voG)3PD;@uu)S1CUJ6{(e4*hLQ0r%;ymYMs7knh#vCTdEjlwO->>RF#dc++y8}5&K@lwz01l?Fdr7I|!S1MR;|}Q|(cSHNJV#sl!Ns=<fI8px4`fd1o{f)PmI^s1W@}i&^JNch|Ec7n?z=7L)YBPsnlaVr{za*_(>U4vBiCk`g<cV`b*))n_;Ursq|-YjuXN%t{7d40g<^WhTLQMlIDUZNAbjrKo1bR4;uBiQU1^J=0L3?a6n*)DeKL{{7WBclO_BJaFU7nR5A49dGwNCR3%3D>1*j{@MJSyq2C?Sj+$bwK<XEK>)So0>&yi>tiUu3vH{!Ai@Fcg8UK?$n%jJzs8W)X)A7JPZCvNVDq|2GBLZGsoBtAigBB^$qW*N2N<6e^)_vr+VglXg!B?JAR8w&Kda82NudL1lH%}>@BwX;?6&Uak4Aq8Bg7A31k4C16@T)@%%^ZdW3RjD&DuyB$PuOv=5r~#+Cgx^-;05k=G&;mA_N=zO(!;Yi4Pv)UB~;P7f#;8<0*&8V~^hXat}LDLg!hCyy+dUHp35X=V%W=P|oV*n#{DysHWzSgAJE6|3OhD;X~l_Wz@{KUOgn?Ds2WOn)j#&M@y`*UpCFAmS%U-JwU-nWiU2KFkMy^GJDx)eE78?Q*Nb5+~LI;cSuQ-u}=6AwP0dY(K^`DL6TGcM}KQEe~MpJe>-Bh(C4w>!3bM+*-hl6NG^NqcGug~)_ho%_u7XYm~@g#bKP7A$9Ws!E2SR-9Nw9c5HPI9l$vaD(%m9`1UrO?NZ!>e7knd&eG1!?AZHKAL<Xx$ekVz!%Jb>eX_dV5Aq>JUjUEUA>g-?|+co(%5e`wfO8mt@p6(1H5oYP<{jEACu2;U(L6`2m_(SeEuTHbG4HtnkNK%IKh24~9{jfqM;L_grR`j`gxTLt@-ILIbOFp$pth>$2`cS<9!_Wy|A0IS$$HX8F_m?1wzn@d?Gmvs;m>oX}X}@GEoT3(*dY_e2oL5h>&WO6hO#4o|obL|n+FHeugZHJ@o#S!eXK=~8c7m<CUx~4HgtIc=Qt-QgVlMv=7BO6~'
COLS = ['ciment','laitier','cendres','eau','adjuvant','gravier','sable','age','resistance']
LABELS = dict(zip(COLS,['Ciment','Laitier','Cendres volantes','Eau','Superplastifiant','Gros granulats','Granulats fins','Âge','Résistance']))
UNITS = {k: 'kg/m³' for k in COLS[:7]} | {'age':'jours','resistance':'MPa','eau_liant':'sans unité'}
LABELS['eau_liant'] = 'Eau / liant (masse)'
ENGLISH = ['Cement','Blast Furnace Slag','Fly Ash','Water','Superplasticizer','Coarse Aggregate','Fine Aggregate','Age','Concrete compressive strength']
TEAL, PURPLE, INK, PEACH = '#006D77', '#6651D5', '#21183A', '#D56A3D'
MAX_ROWS = 10000


def normalize_name(value):
    value = unicodedata.normalize('NFKD', str(value)).encode('ascii','ignore').decode().lower()
    return re.sub(r'[^a-z0-9]', '', value)


def aliases():
    out={}
    for key, en in zip(COLS, ENGLISH):
        choices=[key, LABELS[key], en, key+'_'+UNITS[key], key+'_mpa' if key=='resistance' else key+'_kg_m3']
        for item in choices: out[normalize_name(item)]=key
    for name,key in [('resistancecompression','resistance'),('resistancealacompressionmpa','resistance'),('strength','resistance'),('agejours','age'),('agejour','age'),('cendresvolantes','cendres'),('superplastifiant','adjuvant'),('granulatsfins','sable'),('grosgranulats','gravier')]:
        out[name]=key
    return out


def standardize(frame):
    if not isinstance(frame, pd.DataFrame) or not 1 <= len(frame) <= MAX_ROWS:
        raise ValueError(f'Le tableau doit contenir de 1 à {MAX_ROWS:,} lignes.'.replace(',',' '))
    if len(frame.columns)>50: raise ValueError('Maximum : 50 colonnes à l’import.')
    found={}; ignored=[]; lookup=aliases()
    for original in frame.columns:
        n=normalize_name(original)
        key=lookup.get(n)
        # Reconnaissance des anciens en-têtes du fichier XLS officiel.
        if key is None:
            for en, k in zip(ENGLISH,COLS):
                if n.startswith(normalize_name(en)) and ('component' in n or 'mpa' in n): key=k; break
        if key:
            if key in found: raise ValueError(f'Deux colonnes correspondent à {key}. Renommez ou retirez le doublon.')
            found[key]=original
        else: ignored.append(str(original)[:100])
    if 'resistance' not in found:
        raise ValueError('Colonne requise absente : resistance (MPa). Téléchargez le modèle pour les noms acceptés.')
    data={}
    for key in COLS:
        series=frame[found[key]] if key in found else pd.Series(['']*len(frame),index=frame.index)
        data[key]=series.map(lambda x: '' if pd.isna(x) else str(x).strip()).tolist()
        if any(len(x.encode('utf-8'))>100 for x in data[key]): raise ValueError('Une cellule dépasse 100 octets de texte. Importez seulement des valeurs numériques.')
    return pd.DataFrame(data), ignored


def numbers(raw):
    frame=pd.DataFrame(index=raw.index)
    for key in COLS:
        s=raw[key].astype(str).str.strip().str.replace('\u202f','',regex=False).str.replace('\u00a0','',regex=False).str.replace(' ','',regex=False).str.replace(',','.',regex=False)
        frame[key]=pd.to_numeric(s,errors='coerce').astype(float)
    return frame.replace([np.inf,-np.inf],np.nan)


def clean_data(raw, remove_duplicates=False, complete=False):
    df=numbers(raw)
    reasons=[[] for _ in range(len(df))]
    for key in COLS:
        supplied=raw[key].str.strip().ne('')
        bad=supplied & (df[key].isna() | (df[key]<0) | (df[key]>1e9))
        if key in ['resistance','age','ciment','eau']: bad |= supplied & df[key].le(0)
        if key=='age': bad |= supplied & (df[key].mod(1).fillna(0).ne(0))
        if key=='resistance': bad |= df[key].isna()
        for i in np.flatnonzero(bad.to_numpy()): reasons[i].append(key+' : manquant ou invalide' if key=='resistance' else key+' : valeur invalide')
    if complete:
        for i in np.flatnonzero(df[COLS].isna().any(axis=1).to_numpy()): reasons[i].append('Composition ou âge incomplet')
    invalid=np.array([bool(v) for v in reasons])
    duplicate=df.duplicated(keep='first').to_numpy() & ~invalid
    if remove_duplicates:
        for i in np.flatnonzero(duplicate): reasons[i].append('Doublon complet retiré sur demande')
    retained=np.array([not v for v in reasons])
    clean=df.loc[retained].copy()
    clean.insert(0,'ligne',np.flatnonzero(retained)+1)
    binder=clean[['ciment','laitier','cendres']].sum(axis=1,min_count=3)
    clean['eau_liant']=clean['eau'].div(binder.where(binder>0))
    audit=raw.copy();audit.insert(0,'ligne',np.arange(1,len(raw)+1))
    audit['décision']=['Conservée' if not v else 'Exclue' for v in reasons]
    audit['motif']=[' ; '.join(v) for v in reasons]
    return clean.reset_index(drop=True), audit, dict(raw=len(raw),valid=int(retained.sum()),excluded=int((~retained).sum()),duplicates=int(duplicate.sum()))


def descriptive(values):
    x=pd.Series(values,dtype=float).dropna().to_numpy()
    n=len(x)
    if not n: return dict(n=0,mean=None,median=None,variance=None,std=None,cv=None,min=None,max=None,q05=None,q25=None,q75=None,q95=None)
    sd=float(np.std(x,ddof=1)) if n>1 else None
    mean=float(np.mean(x))
    return dict(n=n,mean=mean,median=float(np.median(x)),variance=sd**2 if sd is not None else None,std=sd,
        cv=100*sd/mean if sd is not None and mean>0 else None,min=float(x.min()),max=float(x.max()),
        q05=float(np.quantile(x,.05)),q25=float(np.quantile(x,.25)),q75=float(np.quantile(x,.75)),q95=float(np.quantile(x,.95)))


def stat_table(frame):
    rows=[]
    for key in COLS+['eau_liant']:
        s=descriptive(frame[key])
        rows.append({'Variable':LABELS[key], 'Unité':UNITS[key], 'n':s['n'], 'Moyenne':s['mean'],'Médiane':s['median'],
            'Variance (n-1)':s['variance'],'Écart-type (n-1)':s['std'],'CV (%)':s['cv'],'Minimum':s['min'],'Maximum':s['max'],'Q05':s['q05'],'Q95':s['q95']})
    return pd.DataFrame(rows)


def threshold_analysis(values, threshold):
    s=descriptive(values);x=np.asarray(values,dtype=float);x=x[np.isfinite(x)]
    if not len(x): return None
    k=int((x<threshold).sum())
    ci=stats.binomtest(k,len(x)).proportion_ci(confidence_level=.95,method='wilson')
    normal=float(stats.norm.cdf(threshold,loc=s['mean'],scale=s['std'])) if s['std'] is not None and s['std']>0 else None
    return dict(count=k,n=len(x),empirical=k/len(x),low=float(ci.low),high=float(ci.high),normal=normal)


def paired(frame,key):
    return frame[[key,'resistance']].dropna()


def relationship(frame,key):
    pair=paired(frame,key);n=len(pair)
    result=dict(n=n,pearson=None,spearman=None,slope=None,intercept=None,r2=None)
    if n<3 or pair[key].nunique()<2 or pair['resistance'].nunique()<2:return result
    x,y=pair[key].to_numpy(),pair['resistance'].to_numpy()
    fit=stats.linregress(x,y)
    result.update(pearson=float(fit.rvalue),spearman=float(stats.spearmanr(x,y).statistic),slope=float(fit.slope),intercept=float(fit.intercept),r2=float(fit.rvalue**2))
    return result


@st.cache_data(max_entries=8,ttl=1800,show_spinner=False)
def bootstrap_mean(values, repeats=2000, seed=42):
    x=np.asarray(values,dtype=float); x=x[np.isfinite(x)]
    if not 2<=len(x)<=MAX_ROWS or not 1000<=int(repeats)<=5000 or not 0<=int(seed)<=2147483647:
        raise ValueError('Bootstrap : 2 à 10 000 valeurs, 1 000 à 5 000 tirages et graine autorisée.')
    rng=np.random.default_rng(int(seed)); means=np.empty(int(repeats))
    for start in range(0,len(means),50):
        size=min(50,len(means)-start)
        means[start:start+size]=x[rng.integers(0,len(x),size=(size,len(x)))].mean(axis=1)
    lo,hi=np.quantile(means,[.025,.975])
    return dict(low=float(lo),high=float(hi),means=means,repeats=int(repeats),seed=int(seed))


def source_bytes():
    data=zlib.decompress(base64.b85decode(EMBEDDED_UCI))
    if hashlib.sha256(data).hexdigest()!=DATA_SHA256: raise ValueError('La copie UCI ne correspond plus à son empreinte attendue.')
    return data


def uci_frame():
    return standardize(pd.read_csv(io.BytesIO(source_bytes())))[0]


def fictional_frame():
    # Jeu volontairement simple, sans prétention de relation physique entre ses valeurs.
    return pd.DataFrame({'ciment':['300']*5,'laitier':['0']*5,'cendres':['0']*5,'eau':['150']*5,'adjuvant':['3']*5,
        'gravier':['1050']*5,'sable':['800']*5,'age':['28']*5,'resistance':['20','25','30','35','40']})


def read_upload(content, name, sheet=0):
    if not content or len(content)>5*1024*1024:raise ValueError('Fichier vide ou supérieur à 5 Mo.')
    if name.lower().endswith('.xlsx'):
        try:
            with zipfile.ZipFile(io.BytesIO(content)) as z:
                if len(z.infolist())>500 or sum(x.file_size for x in z.infolist())>30*1024*1024: raise ValueError('Classeur décompressé trop volumineux.')
            df=pd.read_excel(io.BytesIO(content),engine='openpyxl',sheet_name=sheet,nrows=MAX_ROWS+1,dtype=str)
        except (zipfile.BadZipFile,KeyError,ValueError) as exc:raise ValueError('Classeur XLSX illisible ou trop volumineux : '+str(exc)[:130]) from None
    elif name.lower().endswith('.csv'):
        try:
            txt=content.decode('utf-8-sig')
        except UnicodeDecodeError:
            try: txt=content.decode('cp1252')
            except UnicodeDecodeError:raise ValueError('Encodage non reconnu. Enregistrez le CSV en UTF-8.') from None
        header=txt.splitlines()[0] if txt.splitlines() else ''
        separator=';' if ';' in header else '\t' if '\t' in header else ',' if ',' in header else ';'
        try: df=pd.read_csv(io.StringIO(txt),sep=separator,engine='python',dtype=str,nrows=MAX_ROWS+1)
        except (pd.errors.ParserError,pd.errors.EmptyDataError,ValueError):raise ValueError('CSV illisible. Vérifiez les séparateurs et la ligne d’en-tête.') from None
    else:raise ValueError('Formats acceptés : CSV et XLSX.')
    return standardize(df)


def safe_frame(df):
    out=df.copy()
    for key in out.columns:
        if out[key].dtype==object or pd.api.types.is_string_dtype(out[key].dtype):
            out[key]=out[key].map(lambda x:"'"+x if isinstance(x,str) and x.lstrip().startswith(('=','+','-','@')) else x)
    return out


def csv_bytes(df):
    return safe_frame(df).to_csv(index=False,sep=';',decimal=',',float_format='%.10g').encode('utf-8-sig')


def backup_bytes(raw,meta,rules,filters):
    doc=dict(application=APP_NAME,version=1,metadata=meta,raw=raw.fillna('').to_dict(orient='records'),rules=rules,filters=filters)
    return json.dumps(doc,ensure_ascii=False,allow_nan=False,indent=2).encode('utf-8')


def read_backup(content):
    if len(content)>12*1024*1024:raise ValueError('Sauvegarde supérieure à 12 Mo.')
    try: doc=json.loads(content)
    except (ValueError,UnicodeDecodeError):raise ValueError('JSON illisible.') from None
    if not isinstance(doc,dict) or doc.get('application')!=APP_NAME or type(doc.get('version')) is not int or doc['version']!=1:raise ValueError('Sauvegarde BétonStat version 1 attendue.')
    records=doc.get('raw');meta=doc.get('metadata');rules=doc.get('rules');filters=doc.get('filters')
    if not isinstance(records,list) or not 1<=len(records)<=MAX_ROWS or any(not isinstance(x,dict) or set(x)!=set(COLS) for x in records):raise ValueError('Structure des données incorrecte.')
    if any(not isinstance(v,str) or len(v.encode('utf-8'))>100 for row in records for v in row.values()):raise ValueError('Cellules de sauvegarde non valides.')
    if not isinstance(meta,dict) or set(meta)!=set(['name','kind','citation','retrieved','hash']):raise ValueError('Provenance incorrecte.')
    if any(not isinstance(v,str) or len(v)>2000 for v in meta.values()):raise ValueError('Texte de provenance incorrect.')
    if not isinstance(rules,dict) or set(rules)!=set(['dedup','complete']) or any(type(v)is not bool for v in rules.values()):raise ValueError('Règles de nettoyage incorrectes.')
    if not isinstance(filters,dict) or set(filters)!=set(['age_mode','age_low','age_high','threshold']):raise ValueError('Filtres de sauvegarde incorrects.')
    if filters['age_mode'] not in ['Tous âges','28 jours','Plage d’âge']:raise ValueError('Mode de filtre incorrect.')
    for key in ['age_low','age_high','threshold']:
        v=filters[key]
        if isinstance(v,bool) or not isinstance(v,(int,float)) or not math.isfinite(v) or not 0<=v<=1e9:raise ValueError('Valeur de filtre incorrecte.')
    if filters['age_low']>filters['age_high'] or filters['threshold']>1000:raise ValueError('Plage de filtre incorrecte.')
    raw=pd.DataFrame(records,columns=COLS)
    # Une provenance UCI n’est conservée qu’en présence du tableau original exact.
    if meta['kind']=='UCI documenté' and not raw.equals(uci_frame()):
        meta={**meta,'name':'Import modifié issu d’une sauvegarde','kind':'Origine déclarée','citation':meta['citation']+' | Étiquette UCI non vérifiée sur les données restaurées.'}
    return raw,meta,rules,filters


def fmt(v, digits=2):
    if v is None or pd.isna(v):return 'Non calculable'
    return f'{float(v):,.{digits}f}'.replace(',',' ').replace('.',',')


def source_meta(name,kind,citation='',hashvalue=''):
    return dict(name=name,kind=kind,citation=citation,retrieved=datetime.now(timezone.utc).isoformat(timespec='seconds'),hash=hashvalue)


def activate(raw,meta,rules=None,filters=None):
    st.session_state.bs_raw=raw.copy()
    st.session_state.bs_meta=meta
    st.session_state.bs_revision=st.session_state.get('bs_revision',0)+1
    st.session_state.bs_defaults=dict(rules=rules or dict(dedup=False,complete=False),filters=filters or dict(age_mode='Tous âges',age_low=1.,age_high=365.,threshold=30.))
    for key in ['bs_boot','bs_reports','bs_comparison']:st.session_state.pop(key,None)


def load_uci():
    activate(uci_frame(),source_meta('UCI · 1 030 essais','UCI documenté',SOURCE_CITATION,DATA_SHA256))


def load_demo():
    activate(fictional_frame(),source_meta('Test guidé · 5 essais fictifs','Exemple fictif','Jeu pédagogique construit pour vérifier les calculs.'))


def current():
    return st.session_state.bs_context


def need_data():
    ctx=current()
    if ctx['df'].empty:
        st.info('Aucun essai analysable dans cette sélection. Chargez des données ou élargissez les filtres ; le journal de nettoyage reste consultable dans Données.')
        return None
    return ctx


def context_signature(ctx):
    return hashlib.sha256(json.dumps(dict(raw=ctx['raw'].to_dict(orient='records'),meta=ctx['meta'],rules=ctx['rules'],filters=ctx['filters']),ensure_ascii=False,sort_keys=True).encode()).hexdigest()


def grid(df, caption, height='auto'):
    st.dataframe(df,hide_index=True,width='stretch',height=height,alt=caption)


def callout(title,body,tone='teal'):
    st.html(f'<div class="bs-note {tone}"><strong>{escape(title)}</strong><p>{escape(body)}</p></div>')


def metrics(items):
    st.html('<div class="bs-metrics">'+''.join(f'<div class="bs-metric"><span>{escape(a)}</span><strong>{escape(b)}</strong><small>{escape(c)}</small></div>' for a,b,c in items)+'</div>')


def heading(kicker,title,description):
    st.html(f'<div class="bs-heading"><span>{escape(kicker)}</span><h1>{escape(title)}</h1><p>{escape(description)}</p></div>')


def chart(ch,description,height=310):
    ch=ch.properties(height=height).configure_axis(labelColor='#4B4662',titleColor=INK,gridColor='#E8E5F0',labelFontSize=11,titleFontSize=12).configure_view(strokeWidth=0).configure_legend(labelColor=INK,titleColor=INK)
    st.altair_chart(ch,width='stretch',theme=None,alt=description)


def histogram(df,bins=18):
    return alt.Chart(df).mark_bar(color=PURPLE,cornerRadiusTopLeft=3,cornerRadiusTopRight=3).encode(
        x=alt.X('resistance:Q',bin=alt.Bin(maxbins=bins),title='Résistance (MPa)'),y=alt.Y('count():Q',title='Nombre d’essais'),tooltip=[alt.Tooltip('count():Q',title='Essais')])


def age_groups(df):
    d=df.copy();d['Classe d’âge']=pd.cut(d['age'],bins=[0,7,28,90,np.inf],labels=['1–7 j','8–28 j','29–90 j','> 90 j']).astype(object).fillna('Âge inconnu')
    return d


def start_page_v1():
    ctx=current()
    st.html('<section class="bs-hero"><div><span class="bs-eyebrow">ATELIER D’ANALYSE · GÉNIE CIVIL</span><h1>Le béton, à travers<br>ses données.</h1><p>Explorer les essais. Mesurer la dispersion.<br>Comprendre les relations entre composition et résistance.</p></div><div class="bs-orbit" aria-hidden="true"><span>σ</span><i>MPa</i><b>n</b><em>μ</em></div></section>')
    if ctx['raw'].empty:
        with st.container(horizontal=True,gap='medium'):
            st.button('Charger les 1 030 essais UCI',icon=':material/database:',type='primary',on_click=load_uci,width='stretch',wrap=True)
            st.button('Essayer 5 valeurs fictives',icon=':material/science:',on_click=load_demo,width='stretch',wrap=True)
        st.html('<div class="bs-path"><div><b>01</b><strong>Collecter</strong><span>Source documentée ou import</span></div><div><b>02</b><strong>Examiner</strong><span>Qualité, unités et sous-ensemble</span></div><div><b>03</b><strong>Analyser</strong><span>Statistiques et graphiques</span></div><div><b>04</b><strong>Expliquer</strong><span>Résultats et hypothèses</span></div></div>')
        callout('Votre point de départ','Les données UCI sont une copie documentée d’essais expérimentaux. Le petit jeu fictif sert uniquement à tester les formules. Aucune source n’est mélangée automatiquement à une autre.')
        st.markdown(f'[Consulter la fiche UCI]({SOURCE_URL}) · [DOI et attribution](https://doi.org/10.24432/C5PK67)')
        return
    c=need_data()
    if c is None:return
    df=c['df'];s=descriptive(df['resistance']);t=threshold_analysis(df['resistance'],c['filters']['threshold'])
    metrics([('Essais sélectionnés',fmt(s['n'],0),f"{c['quality']['valid']} essais après nettoyage"),('Résistance moyenne',fmt(s['mean'])+' MPa','Moyenne de la sélection'),('Dispersion',fmt(s['std'])+' MPa' if s['std'] is not None else 'Indisponible','Écart-type d’échantillon (n−1)'),('Sous le seuil',fmt(100*t['empirical'],1)+' %',f"Seuil exploratoire : {fmt(c['filters']['threshold'])} MPa")])
    with st.container(horizontal=True,gap='medium'):
        st.page_link(st.session_state.bs_pages['distributions'],label='Explorer les distributions',icon=':material/stacked_bar_chart:',width='stretch')
        st.page_link(st.session_state.bs_pages['relations'],label='Étudier les relations',icon=':material/scatter_plot:',width='stretch')
    a,b=st.columns([1.3,1])
    with a:
        st.subheader('Profil des résistances')
        chart(histogram(df),'Répartition des résistances du sous-ensemble sélectionné')
    with b:
        st.subheader('Résistance et classes d’âge')
        d=age_groups(df)
        chart(alt.Chart(d).mark_boxplot(color=TEAL,extent=1.5).encode(x=alt.X('Classe d’âge:N',title=None,sort=['1–7 j','8–28 j','29–90 j','> 90 j','Âge inconnu']),y=alt.Y('resistance:Q',title='Résistance (MPa)')),'Dispersion des résistances par classe d’âge')
    callout('Une comparaison utile','Comparer des résistances à des âges différents peut masquer des différences de formulation. Essayez le filtre « 28 jours » avant d’interpréter les corrélations. Les données ne représentent pas un chantier camerounais identifié.')


def data_page():
    ctx=current();revision=st.session_state.bs_revision
    heading('01 / DONNÉES','Une base lisible, une origine explicite.','Importer, examiner et corriger avant de calculer.')
    with st.container(horizontal=True,gap='medium'):
        st.button('Charger la source UCI',on_click=load_uci,icon=':material/database:',type='primary',width='stretch',wrap=True)
        st.button('Charger le test fictif',on_click=load_demo,icon=':material/science:',width='stretch',wrap=True)
    st.caption('Ces boutons remplacent le jeu actif dans cette session. Exportez une sauvegarde avant de changer de source.')
    with st.expander('Importer un CSV ou un classeur XLSX',expanded=ctx['raw'].empty):
        st.download_button('Télécharger le modèle CSV',csv_bytes(fictional_frame().iloc[:0]),'betonstat-modele.csv','text/csv',on_click='ignore',icon=':material/download:')
        uploaded=st.file_uploader('Votre fichier (CSV / XLSX, 5 Mo maximum)',type=['csv','xlsx'],max_upload_size=5,key='bs_upload')
        st.caption('Colonne minimale : resistance, en MPa. Colonnes facultatives : ciment, laitier, cendres, eau, adjuvant, gravier, sable (kg/m³), age (jours entiers). XLSX : première feuille uniquement.')
        if uploaded:
            try:
                preview,ignored=read_upload(uploaded.getvalue(),uploaded.name)
                grid(preview.head(8),'Aperçu des colonnes reconnues du fichier')
                if ignored:st.info('Colonnes ignorées : '+', '.join(ignored))
                label=st.text_input('Nom ou description de la source',uploaded.name,max_chars=200,key='bs_upload_label')
                confirmed=st.checkbox('Remplacer le jeu actif par ce fichier',key='bs_confirm_import')
                if st.button('Utiliser ce fichier',disabled=not confirmed,type='primary'):
                    activate(preview,source_meta(label.strip() or 'Import utilisateur','Origine déclarée','Données fournies par l’utilisateur ; provenance non vérifiée.',hashlib.sha256(uploaded.getvalue()).hexdigest()));st.rerun()
            except (ValueError,TypeError,OSError,OverflowError) as exc:st.error(str(exc))
    with st.expander('Saisir des essais ou modifier le tableau',expanded=False):
        st.caption('Chaque ligne représente un essai. Les valeurs vides restent manquantes ; elles ne sont jamais remplacées automatiquement par zéro. Tout changement doit être enregistré.')
        mode=st.selectbox('Point de départ',['Tableau actif','Nouveau tableau vide'],key='bs_edit_mode')
        initial=ctx['raw'].copy() if mode=='Tableau actif' and not ctx['raw'].empty else pd.DataFrame([{k:'' for k in COLS} for _ in range(5)])
        if len(initial)>1500:st.info('Pour modifier plus de 1 500 lignes, exportez le CSV, modifiez-le puis réimportez-le.')
        else:
            with st.form('bs_editor_form_'+str(revision)+'_'+mode):
                edited=st.data_editor(initial,num_rows='dynamic',hide_index=True,key='bs_editor_'+str(revision)+'_'+mode,alt='Essais à saisir ou corriger',height=285)
                confirmed=st.checkbox('Confirmer le remplacement du jeu actif par ce tableau')
                run=st.form_submit_button('Enregistrer le tableau',type='primary')
            if run:
                try:
                    if not confirmed:raise ValueError('Confirmez le remplacement du jeu actif.')
                    edited=edited.fillna('');edited=edited.loc[edited.astype(str).apply(lambda x:x.str.strip()).ne('').any(axis=1)]
                    raw,_=standardize(edited)
                    kind='Exemple fictif' if mode=='Tableau actif' and ctx['meta']['kind']=='Exemple fictif' else 'Saisie ou correction utilisateur'
                    activate(raw,source_meta('Tableau saisi ou corrigé',kind,'Valeurs enregistrées par l’utilisateur, sans validation du protocole expérimental.'));st.rerun()
                except (ValueError,TypeError) as exc:st.error(str(exc))
    if ctx['raw'].empty:return
    st.subheader('Qualité et traçabilité')
    q=ctx['quality'];metrics([('Lignes importées',fmt(q['raw'],0),'Avant nettoyage'),('Lignes conservées',fmt(q['valid'],0),'Avant filtre d’âge'),('Lignes exclues',fmt(q['excluded'],0),'Raisons détaillées ci-dessous'),('Doublons complets',fmt(q['duplicates'],0),'Après examen des valeurs invalides')])
    st.caption('Les doublons sont conservés par défaut : une ligne répétée peut représenter une répétition expérimentale légitime. La suppression se décide dans le menu latéral.')
    view=st.segmented_control('Tableau à consulter',['Données sélectionnées','Journal de nettoyage','Données brutes'],default='Données sélectionnées')
    frame=ctx['audit'] if view=='Journal de nettoyage' else ctx['raw'] if view=='Données brutes' else ctx['df']
    grid(frame,'Tableau de données et décisions de nettoyage',height=400)
    with st.container(horizontal=True):
        st.download_button('Données sélectionnées CSV',csv_bytes(ctx['df']),'betonstat-donnees.csv','text/csv',on_click='ignore',width='stretch',wrap=True)
        st.download_button('Journal de nettoyage CSV',csv_bytes(ctx['audit']),'betonstat-nettoyage.csv','text/csv',on_click='ignore',width='stretch',wrap=True)
    with st.expander('Dictionnaire des variables et provenance'):
        grid(pd.DataFrame([{'Colonne':k,'Variable':LABELS[k],'Unité':UNITS[k]} for k in COLS]),'Dictionnaire des variables')
        st.markdown('**Source active :** '+escape(ctx['meta']['name']))
        st.caption(ctx['meta']['citation'])
        st.code('SHA-256 de la source : '+ctx['meta']['hash'] if ctx['meta']['hash'] else 'Pas d’empreinte de fichier pour une saisie manuelle.')
        st.markdown(f'[Fiche UCI]({SOURCE_URL}) · [CSV officiel]({DATA_URL}) · [Licence CC BY 4.0](https://creativecommons.org/licenses/by/4.0/)')


def statistics_page():
    heading('02 / STATISTIQUES','Mesurer le niveau et la dispersion.','Toutes les valeurs sont recalculées sur le sous-ensemble actif, sans remplacer les données manquantes.')
    ctx=need_data()
    if ctx is None:return
    df=ctx['df'];s=descriptive(df['resistance'])
    metrics([('Moyenne',fmt(s['mean'])+' MPa','Somme / nombre d’essais'),('Médiane',fmt(s['median'])+' MPa','Valeur centrale'),('Variance',fmt(s['variance'])+' MPa²' if s['variance'] is not None else 'Indisponible','Diviseur n−1'),('Coefficient de variation',fmt(s['cv'])+' %' if s['cv'] is not None else 'Indisponible','Écart-type / moyenne × 100')])
    grid(stat_table(df).round(5),'Statistiques descriptives de toutes les variables',height=400)
    st.caption('Variance et écart-type : estimation d’échantillon (n−1), définis seulement à partir de 2 observations. Chaque variable utilise ses propres valeurs renseignées ; les effectifs n peuvent différer. Quantiles interpolés linéairement.')
    guided_calculations(ctx)
    with st.expander('Voir les formules et un calcul à la main'):
        st.latex(r'\bar{x}=\frac{1}{n}\sum x_i,\quad s^2=\frac{\sum(x_i-\bar{x})^2}{n-1},\quad CV=100\frac{s}{\bar{x}}')
        st.markdown('Pour **20, 25, 30, 35, 40 MPa** : moyenne = médiane = **30 MPa** ; somme des carrés des écarts = **250 MPa²** ; variance = **62,5 MPa²** ; écart-type ≈ **7,9057 MPa** ; CV ≈ **26,3523 %**.')
    st.subheader('Incertitude sur la moyenne : bootstrap')
    st.caption('On rééchantillonne les essais avec remise pour examiner la variabilité de la moyenne. Cette méthode suppose des observations comparables et indépendantes ; elle ne corrige ni un échantillon biaisé ni des répétitions liées.')
    with st.form('bs_boot_form'):
        a,b=st.columns(2);repeats=a.select_slider('Rééchantillonnages',[1000,2000,5000],value=2000);seed=b.number_input('Graine du bootstrap',0,2147483647,42)
        run=st.form_submit_button('Estimer l’intervalle de la moyenne',type='primary',disabled=len(df)<2)
    sig=context_signature(ctx)
    if run:
        with st.spinner('Rééchantillonnages en cours…'):
            st.session_state.bs_boot=(sig,bootstrap_mean(tuple(df['resistance']),repeats,seed))
    boot=st.session_state.get('bs_boot')
    if boot and boot[0]==sig:
        r=boot[1];callout('Intervalle bootstrap percentile à 95 %',f"Moyenne estimée entre {fmt(r['low'])} et {fmt(r['high'])} MPa. {r['repeats']} rééchantillonnages, graine {r['seed']}. Cet intervalle concerne la moyenne, pas 95 % des éprouvettes.")
        chart(alt.Chart(pd.DataFrame({'moyenne':r['means']})).mark_bar(color=TEAL).encode(x=alt.X('moyenne:Q',bin=alt.Bin(maxbins=25),title='Moyenne rééchantillonnée (MPa)'),y=alt.Y('count():Q',title='Rééchantillonnages')),'Variabilité simulée de la moyenne par bootstrap',260)


def distributions_page():
    heading('03 / DISTRIBUTIONS','Lire une fréquence, discuter un modèle.','La fréquence observée et la probabilité issue d’une loi ajustée sont présentées séparément.')
    ctx=need_data()
    if ctx is None:return
    df=ctx['df'];x=df['resistance'].to_numpy();s=descriptive(x);threshold=ctx['filters']['threshold'];t=threshold_analysis(x,threshold)
    metrics([('Essais sous le seuil',str(t['count'])+' / '+str(t['n']),f'Condition stricte : R < {fmt(threshold)} MPa'),('Fréquence observée',fmt(100*t['empirical'],1)+' %','Sur la sélection active'),('Intervalle de Wilson',fmt(100*t['low'],1)+' – '+fmt(100*t['high'],1)+' %','95 % ; hypothèse d’essais indépendants'),('Modèle normal',fmt(100*t['normal'],1)+' %' if t['normal'] is not None else 'Non ajustable','Probabilité théorique conditionnelle')])
    bins=st.slider('Nombre maximal de classes de l’histogramme',5,40,18)
    rule=alt.Chart(pd.DataFrame({'seuil':[threshold]})).mark_rule(color=PEACH,strokeDash=[6,3],strokeWidth=2).encode(x='seuil:Q')
    chart(histogram(df,bins)+rule,'Histogramme des résistances et seuil exploratoire')
    a,b=st.columns(2)
    with a:
        st.subheader('Fréquence cumulée empirique')
        vals,counts=np.unique(np.sort(x),return_counts=True)
        ecdf=pd.DataFrame({'Résistance (MPa)':vals,'Part inférieure ou égale':np.cumsum(counts)/len(x)})
        chart(alt.Chart(ecdf).mark_line(interpolate='step-after',color=TEAL,point=True).encode(x='Résistance (MPa):Q',y=alt.Y('Part inférieure ou égale:Q',axis=alt.Axis(format='%')),tooltip=list(ecdf.columns)),'Part des observations inférieures ou égales à chaque résistance')
    with b:
        st.subheader('Diagramme quantile-quantile normal')
        if len(x)>1 and s['std']>0:
            expected=stats.norm.ppf((np.arange(1,len(x)+1)-.5)/len(x))*s['std']+s['mean']
            qq=pd.DataFrame({'Quantile normal (MPa)':expected,'Résistance observée (MPa)':np.sort(x)})
            reference=alt.Chart(pd.DataFrame({'x':[float(expected.min()),float(expected.max())]})).mark_line(color=PEACH,strokeDash=[4,3]).encode(x='x:Q',y='x:Q')
            chart(alt.Chart(qq).mark_circle(color=PURPLE,opacity=.65,size=30).encode(x='Quantile normal (MPa):Q',y='Résistance observée (MPa):Q')+reference,'Comparaison des quantiles empiriques et des quantiles d’un modèle normal')
        else:st.info('Deux valeurs et une dispersion non nulle sont nécessaires au diagramme normal.')
    callout('Le modèle normal est une hypothèse','La loi normale utilise la moyenne et l’écart-type d’échantillon affichés. La courbe cumulée empirique compte R ≤ x, tandis que le seuil utilise R < seuil. Un écart au diagramme normal invite à discuter le modèle ; il ne constitue pas une décision de conformité. La loi normale peut attribuer une masse à des résistances négatives.','purple')
    if s['std'] is not None and s['std']>0:
        tail=float(stats.norm.cdf(0,loc=s['mean'],scale=s['std']))
        st.caption(f'Masse théorique normale sous 0 MPa : {fmt(100*tail,3)} %. Aucun écrêtage silencieux ni troncature n’est appliqué.')
    q1,q3=s['q25'],s['q75'];mask=(x<q1-1.5*(q3-q1))|(x>q3+1.5*(q3-q1))
    st.caption(f'{int(mask.sum())} observation(s) au-delà des bornes Q1 − 1,5 IQR / Q3 + 1,5 IQR. Ce repérage descriptif ne supprime aucune ligne et ne prouve pas une erreur d’essai.')


def relations_page():
    heading('04 / RELATIONS','Repérer les associations, sans surinterpréter.','Une corrélation décrit un lien statistique dans ce jeu, pas l’effet isolé d’un ingrédient.')
    ctx=need_data()
    if ctx is None:return
    df=ctx['df']; keys=[k for k in COLS[:-1]+['eau_liant'] if df[k].notna().sum()>=2]
    if not keys:st.info('Ajoutez une variable de composition ou l’âge pour étudier une relation avec la résistance.');return
    key=st.selectbox('Variable à comparer à la résistance',keys,format_func=lambda x:LABELS[x]+' ('+UNITS[x]+')',key='bs_relation_variable')
    r=relationship(df,key);d=paired(df,key)
    metrics([('Paires complètes',fmt(r['n'],0),'Valeurs renseignées pour les deux variables'),('Pearson r',fmt(r['pearson'],3),'Association linéaire'),('Spearman ρ',fmt(r['spearman'],3),'Association des rangs'),('R² linéaire',fmt(r['r2'],3),'Ajustement descriptif sur les mêmes essais')])
    ch=alt.Chart(d).mark_circle(size=55,color=TEAL,opacity=.65).encode(x=alt.X(key+':Q',title=LABELS[key]+' ('+UNITS[key]+')',scale=alt.Scale(zero=False)),y=alt.Y('resistance:Q',title='Résistance (MPa)'),tooltip=[alt.Tooltip(key+':Q',title=LABELS[key]),alt.Tooltip('resistance:Q',title='Résistance (MPa)')]).interactive()
    if r['slope'] is not None:
        xs=np.array([d[key].min(),d[key].max()]);line=pd.DataFrame({key:xs,'resistance':r['intercept']+r['slope']*xs})
        ch=ch+alt.Chart(line).mark_line(color=PEACH,strokeWidth=2).encode(x=key+':Q',y='resistance:Q')
    chart(ch,'Relation entre '+LABELS[key]+' et la résistance',380)
    if r['slope'] is not None:st.caption(f"Droite descriptive : résistance = {fmt(r['intercept'],4)} + ({fmt(r['slope'],4)}) × {LABELS[key]}. Ajustement interne uniquement ; aucune validation prédictive hors échantillon.")
    else:st.info('Corrélations et droite non calculables : au moins 3 paires et de la variabilité dans les deux variables sont nécessaires.')
    st.subheader('Matrice des associations')
    method=st.segmented_control('Méthode de corrélation',['Pearson','Spearman'],default='Pearson')
    vars=[k for k in COLS+['eau_liant'] if df[k].notna().sum()>=3 and df[k].nunique()>1]
    if len(vars)>=2:
        corr=df[vars].corr(method=method.lower(),min_periods=3)
        counts=df[vars].notna().astype(int).T.dot(df[vars].notna().astype(int))
        cells=pd.DataFrame([{'Variable X':LABELS[a],'Variable Y':LABELS[b],'Corrélation':corr.loc[a,b],'Paires':int(counts.loc[a,b])} for a in vars for b in vars])
        chart(alt.Chart(cells).mark_rect().encode(x=alt.X('Variable X:N',title=None,axis=alt.Axis(labelAngle=-40,labelLimit=160)),y=alt.Y('Variable Y:N',title=None),color=alt.Color('Corrélation:Q',scale=alt.Scale(domain=[-1,0,1],range=[PEACH,'#F7F4FA',TEAL])),tooltip=['Variable X:N','Variable Y:N',alt.Tooltip('Corrélation:Q',format='.3f'),'Paires:Q']),'Corrélations calculées sur les paires de valeurs disponibles',410)
        with st.expander('Voir la matrice et les effectifs de chaque paire'):
            grid(corr.rename(index=LABELS,columns=LABELS).reset_index(),'Coefficients de corrélation');grid(counts.rename(index=LABELS,columns=LABELS).reset_index(),'Nombre de paires par coefficient')
    callout('Variables liées et âges mélangés','Le ciment, l’eau et les additions peuvent varier ensemble. Une association observée ne démontre pas la causalité. Filtrer à 28 jours réduit le mélange des âges, sans contrôler tous les autres facteurs. Aucune règle de formulation ou de dimensionnement n’est déduite automatiquement.','purple')
    st.caption('Eau / liant = eau / (ciment + laitier + cendres), en masses. Le rapport est calculé uniquement si les trois constituants et l’eau sont renseignés et si le liant total est positif. Il ne s’agit pas d’un liant équivalent normatif.')


def xlsx_bytes(ctx):
    wb=Workbook();wb.remove(wb.active)
    meta=pd.DataFrame([{'Rubrique':k,'Valeur':str(v)} for k,v in {**ctx['meta'],**ctx['rules'],**ctx['filters']}.items()])
    for title,frame in [('Selection',ctx['df']),('Statistiques',stat_table(ctx['df'])),('Nettoyage',ctx['audit']),('Brut',ctx['raw']),('Provenance',meta)]:
        ws=wb.create_sheet(title)
        ws.append([str(v) for v in frame.columns])
        for record in frame.itertuples(index=False,name=None):
            ws.append([None if pd.isna(v) else v.item() if isinstance(v,np.generic) else v for v in record])
            for cell in ws[ws.max_row]:
                if isinstance(cell.value,str):cell.data_type='s'
        ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
        for col in ws.columns:ws.column_dimensions[col[0].column_letter].width=min(50,max(14,len(str(col[0].value))+3))
    bio=io.BytesIO();wb.save(bio);return bio.getvalue()


def pdf_bytes(ctx,options=None):
    from reportlab.platypus import PageBreak
    options=options or default_report_options(ctx)
    # Polices Unicode fournies par ReportLab : aucune dépendance aux polices de Windows.
    import reportlab
    font_dir=Path(reportlab.__file__).parent/'fonts'
    if 'BSReport' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('BSReport',str(font_dir/'Vera.ttf')))
        pdfmetrics.registerFont(TTFont('BSReport-Bold',str(font_dir/'VeraBd.ttf')))
        pdfmetrics.registerFontFamily('BSReport',normal='BSReport',bold='BSReport-Bold',italic='BSReport',boldItalic='BSReport-Bold')
    styles=getSampleStyleSheet()
    for name in styles.byName:
        styles[name].fontName='BSReport';styles[name].fontSize=9;styles[name].leading=13
    styles['Title'].fontSize=23;styles['Title'].leading=28;styles['Title'].textColor=colors.HexColor(TEAL)
    styles['Heading2'].fontName='BSReport-Bold';styles['Heading2'].fontSize=12;styles['Heading2'].leading=16
    small=ParagraphStyle('table',fontName='BSReport',fontSize=8,leading=11)
    out=io.BytesIO();items=[]
    def para(value,style='BodyText'):items.append(Paragraph(escape(str(value)),styles[style]));items.append(Spacer(1,5))
    def tb(headers,rows,widths):
        data=[[Paragraph(escape(str(v)),small) for v in row] for row in [headers]+rows]
        t=Table(data,colWidths=widths,repeatRows=1,hAlign='LEFT')
        t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#E7F2F1')),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#F6F5FA')]),('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),7),('RIGHTPADDING',(0,0),(-1,-1),7),('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),6)]));items.extend([t,Spacer(1,10)])
    para('BétonStat','Title')
    para('LABORATOIRE INTERACTIF · RAPPORT UNIVERSITAIRE','Heading2')
    items.append(Spacer(1,30))
    para(options['title'],'Title')
    para(ACADEMIC_TITLE)
    items.append(Spacer(1,25))
    tb(['Présentation','Informations'],[['Étudiant / groupe',options['student'] or 'Non renseigné'],['Établissement',options['institution'] or 'Non renseigné'],['Date du rapport',options['date']],['Source',ctx['meta']['name']],['Essais sélectionnés',str(len(ctx['df']))]],[140,345])
    para('Document exploratoire : données, statistiques, graphiques et interprétation. Les informations de présentation sont déclarées par l’utilisateur.')
    items.append(PageBreak())
    para('BétonStat | Résultats statistiques','Title');para(ACADEMIC_TITLE)
    para('Source et sélection','Heading2');para(ctx['meta']['name']+' | '+ctx['meta']['kind']);para(ctx['meta']['citation'])
    para('Généré le '+datetime.now(timezone.utc).strftime('%Y-%m-%d à %H:%M UTC')+'. Jeu et filtres figés au moment de cet export.')
    f=ctx['filters'];q=ctx['quality']
    para(f"Lignes brutes : {q['raw']} ; après nettoyage : {q['valid']} ; exclues : {q['excluded']} ; sélection : {len(ctx['df'])}.")
    para(f"Filtre : {f['age_mode']} ; plage paramétrée {fmt(f['age_low'],0)}–{fmt(f['age_high'],0)} jours. Doublons retirés : {'oui' if ctx['rules']['dedup'] else 'non'}. Composition complète exigée : {'oui' if ctx['rules']['complete'] else 'non'}.")
    para('Statistiques descriptives','Heading2')
    rows=[]
    for k in COLS+['eau_liant']:
        s=descriptive(ctx['df'][k]);rows.append([LABELS[k]+' ('+UNITS[k]+')',str(s['n']),fmt(s['mean']),fmt(s['median']),fmt(s['std']),fmt(s['cv'])])
    tb(['Variable','n','Moyenne','Médiane','Écart-type','CV (%)'],rows,[150,46,73,73,75,68])
    para('Seuil exploratoire et modèle','Heading2')
    t=threshold_analysis(ctx['df']['resistance'],f['threshold'])
    if t:
        para(f"Sous {fmt(f['threshold'])} MPa (strictement) : {t['count']}/{t['n']}, soit {fmt(100*t['empirical'])} %. Intervalle de Wilson à 95 % : {fmt(100*t['low'])} à {fmt(100*t['high'])} %.")
        para('Probabilité sous loi normale ajustée : '+(fmt(100*t['normal'])+' %.' if t['normal'] is not None else 'non calculable.'))
    else:para('Sélection vide : aucun résultat statistique calculable.')
    items.append(PageBreak())
    para('Relations exploratoires','Heading2')
    rels=[]
    for k in COLS[:-1]+['eau_liant']:
        r=relationship(ctx['df'],k)
        if r['pearson'] is not None:rels.append([LABELS[k],str(r['n']),fmt(r['pearson'],3),fmt(r['spearman'],3),fmt(r['r2'],3)])
    if rels:tb(['Variable / résistance','Paires','Pearson','Spearman','R²'],rels,[185,55,80,80,71])
    else:para('Pas assez de paires variables pour estimer une relation.')
    para('Méthodes et limites','Heading2')
    para('Variance et écart-type avec diviseur n-1. Quantiles interpolés linéairement. Données manquantes non imputées. Corrélations sur paires complètes, au moins 3, sans preuve de causalité. Les statistiques dépendent de la composition et des âges sélectionnés.')
    para('Le seuil est choisi par l’utilisateur, sans critère normatif intégré. Les intervalles supposent des essais indépendants et comparables. Le modèle normal peut donner une probabilité à des résistances négatives. Ce rapport ne constitue ni une attestation de conformité ni un dimensionnement.')
    para('La procédure bootstrap est interactive dans Statistiques ; elle n’est pas incluse dans ce rapport. Les graphiques choisis figurent dans les pages suivantes ; les tables détaillées sont dans le classeur Excel.')
    para('Traçabilité','Heading2');para('Source officielle : '+SOURCE_URL);para('CSV : '+DATA_URL);para('Empreinte source : '+(ctx['meta']['hash'] or 'saisie manuelle, sans fichier source'));para('Copie UCI récupérée le '+RETRIEVED+' ; citation et licence ci-dessus.')
    report_additions(items,ctx,options,styles,para,tb)
    def footer(c,doc):
        c.setFont('BSReport',8);c.setFillColor(colors.HexColor(INK));c.drawString(36,23,'BétonStat · Analyse exploratoire');c.drawRightString(A4[0]-36,23,str(doc.page))
    SimpleDocTemplate(out,pagesize=A4,leftMargin=36,rightMargin=36,topMargin=35,bottomMargin=45,title='BétonStat - Rapport statistique',author='BétonStat').build(items,onFirstPage=footer,onLaterPages=footer)
    return out.getvalue()


def methods_page():
    heading('08 / MÉTHODES','Comprendre, tester, défendre.','Un parcours de démonstration et les limites à expliquer dans un projet universitaire.')
    st.markdown('**Titre académique**  \n'+ACADEMIC_TITLE)
    st.subheader('Votre démonstration en cinq étapes')
    st.markdown('1. **Données** : chargez le test fictif, soit 20, 25, 30, 35 et 40 MPa à 28 jours.\n2. **Statistiques** : retrouvez moyenne et médiane 30 MPa, variance 62,5 MPa², écart-type 7,9057 MPa et CV 26,3523 %.\n3. **Distributions** : au seuil de 30 MPa, 2 valeurs sur 5 sont strictement en dessous : 40 %. Le modèle normal ajusté donne 50 %.\n4. **Relations** : dans ce petit jeu, la composition et l’âge sont constants. Les corrélations ne doivent donc pas être annoncées comme calculables. Chargez ensuite UCI pour explorer les associations.\n5. **Rapports** : exportez le PDF, Excel et une sauvegarde JSON, puis restaurez cette sauvegarde.')
    with st.container(horizontal=True):
        st.button('Charger le test de calcul',on_click=load_demo,icon=':material/science:',type='primary',width='stretch',wrap=True)
        st.button('Passer aux essais UCI',on_click=load_uci,icon=':material/database:',width='stretch',wrap=True)
    with st.expander('Contrôles sur les données UCI intégrées',expanded=True):
        st.markdown('**1 030 lignes et 9 variables**, avant nettoyage optionnel. La copie CSV utilisée donne une moyenne de résistance de **35,817961 MPa**, une médiane de **34,445 MPa** et un écart-type d’échantillon de **16,705742 MPa**. Elle contient **25 doublons complets** ; ils restent présents par défaut. Les retirer change les statistiques : ce choix doit être indiqué.')
        st.caption('Ces repères concernent le CSV officiel cité, tous âges confondus, sans exclusion des doublons. Le fichier XLS historique peut conserver davantage de décimales : documentez la version analysée.')
    with st.expander('Lire les limites du modèle'):
        st.markdown('**Petits échantillons :** une moyenne n’explique pas toute la variabilité. Avec une observation, variance et écart-type sont indéfinis.\n\n**Âge :** l’application ne convertit pas automatiquement un essai vers 28 jours. Le filtre 28 jours conserve seulement les essais mesurés à cet âge.\n\n**Composition :** les formulations ne sont pas un plan expérimental contrôlé. L’eau et le ciment peuvent varier avec d’autres ingrédients. Une corrélation ne suffit pas à établir une influence causale.\n\n**Indépendance :** les identifiants de lots ne sont pas fournis dans le schéma retenu ; les intervalles de Wilson et le bootstrap ne corrigent pas une dépendance entre essais.\n\n**Valeurs extrêmes :** aucune valeur n’est supprimée parce qu’elle est simplement éloignée de la moyenne.\n\n**Conformité :** ni résistance caractéristique normative ni réception de lot ne sont certifiées. Le seuil est un repère exploratoire.\n\n**Sources :** le jeu UCI n’est pas une campagne d’essais locale identifiée. La provenance des fichiers de l’utilisateur reste déclarative.')
    with st.expander('Utiliser le logiciel sur Windows'):
        st.code('python -m pip install -r requirement.text\npython -m streamlit run app.py',language='powershell')
        st.caption('Pour Community Cloud, préparez une copie de déploiement dont le manifeste est nommé requirements.txt. Conservez requirement.text dans le dossier local livré.')
    st.markdown(f'**Référence du jeu expérimental :** {SOURCE_CITATION}  \n[Fiche et variables UCI]({SOURCE_URL}) · [DOI](https://doi.org/10.24432/C5PK67) · [CSV utilisé]({DATA_URL})')


# Fonctions intégrées dans app.py par upgrade_betonstat.py ; fichier de travail hors livraison.
MATERIAL_COLORS = dict(zip(COLS[:7], ['#6651D5','#006D77','#B15C35','#2878A0','#9F3873','#72634D','#9A730D']))
GROUP_COLORS = {'Groupe A': TEAL, 'Groupe B': PURPLE}


def insight_messages(ctx):
    d=ctx['df'];q=ctx['quality'];s=descriptive(d.resistance)
    messages=[]
    if q['excluded']:messages.append(f"{q['excluded']} ligne(s) exclue(s) : consultez les motifs dans Données.")
    if d.age.nunique()>1:messages.append('Plusieurs âges sont présents. Comparez de préférence des essais à un âge commun, par exemple 28 jours.')
    if d.age.isna().any():messages.append(f"L’âge manque pour {int(d.age.isna().sum())} essai(s) ; leur comparabilité reste incertaine.")
    if len(d)<20:messages.append('Ce petit échantillon permet une démonstration ; les conclusions et les intervalles restent fragiles.')
    if s['std']==0:messages.append('Toutes les résistances sont identiques : une corrélation ou une loi normale de dispersion non nulle ne peut pas être estimée.')
    if q['duplicates'] and not ctx['rules']['dedup']:messages.append(f"{q['duplicates']} répétition(s) complète(s) sont conservées. Sans identifiants de lots, leur origine reste inconnue.")
    if not messages:messages.append('Les valeurs sont analysables. Vérifiez encore leur provenance et l’indépendance des essais avant de généraliser.')
    return messages


def auto_conclusion(ctx):
    d=ctx['df'];s=descriptive(d.resistance)
    if not len(d):return 'La sélection ne contient aucun essai analysable. Aucun résultat statistique ne peut être conclu ; les filtres et les exclusions doivent être vérifiés.'
    t=threshold_analysis(d.resistance,ctx['filters']['threshold'])
    text=f"La sélection contient {s['n']} essai(s), avec une résistance moyenne de {fmt(s['mean'])} MPa et une médiane de {fmt(s['median'])} MPa. "
    if s['std'] is not None:text+=f"L’écart-type d’échantillon est de {fmt(s['std'])} MPa et le coefficient de variation de {fmt(s['cv'])} %. "
    text+=f"{t['count']} essai(s), soit {fmt(t['empirical']*100)} %, sont strictement sous le seuil exploratoire de {fmt(ctx['filters']['threshold'])} MPa. "
    if d.age.nunique()>1:text+='Ces essais concernent plusieurs âges ; cette moyenne globale ne représente pas une résistance de référence à 28 jours. '
    return text+'Ces résultats décrivent ce jeu de données. Ils ne démontrent pas de causalité et ne constituent pas une validation réglementaire du béton.'


def sparkline(values):
    x=np.asarray(values,dtype=float);x=x[np.isfinite(x)]
    if not len(x):return ''
    counts,_=np.histogram(x,bins=12);maximum=max(int(counts.max()),1)
    bars=''.join(f'<span style="height:{max(2,36*v/maximum):.2f}px"></span>' for v in counts)
    return '<div class="bs-spark" role="img" aria-label="Mini-histogramme des résistances">'+bars+'</div>'


def overview_v2():
    ctx=current()
    if ctx['raw'].empty:
        start_page_v1()
        return
    c=need_data()
    if c is None:return
    d=c['df'];s=descriptive(d.resistance);t=threshold_analysis(d.resistance,c['filters']['threshold'])
    st.html('<section class="bs-hero bs-hero-live"><div><span class="bs-eyebrow">BÉTONSTAT · LABORATOIRE INTERACTIF</span><h1>Vue d’ensemble<br>des essais.</h1><p>Une lecture claire de vos données, de leur variabilité et de leurs limites.</p></div><div class="bs-hero-stat"><small>SÉLECTION ACTIVE</small><strong>'+fmt(len(d),0)+'</strong><span>essais à explorer</span>'+sparkline(d.resistance)+'</div></section>')
    metrics([('Essais sélectionnés',fmt(s['n'],0),f"{c['quality']['valid']} essais après nettoyage"),('Résistance moyenne',fmt(s['mean'])+' MPa','Valeur observée'),('Dispersion',fmt(s['std'])+' MPa' if s['std'] is not None else 'Indisponible','Écart-type d’échantillon (n−1)'),('Sous le seuil',fmt(100*t['empirical'],1)+' %',f"R < {fmt(c['filters']['threshold'])} MPa ; repère exploratoire")])
    missing=int(d[COLS].isna().sum().sum());ages=d.age.dropna()
    age_text=f"{fmt(ages.min(),0)} à {fmt(ages.max(),0)} jours" if len(ages) else 'Âges non renseignés'
    st.html('<div class="bs-zones"><article class="bs-zone"><span class="bs-step">01 / MES DONNÉES</span><h3>Une base à examiner</h3><p><b>'+escape(c['meta']['kind'])+'</b><br>'+escape(age_text)+'<br>'+str(missing)+' cellule(s) manquante(s) dans la sélection.</p><small>'+str(c['quality']['excluded'])+' ligne(s) exclue(s) avant filtrage.</small></article><article class="bs-zone"><span class="bs-step">02 / CE QUE J’OBSERVE</span><h3>'+fmt(s['mean'])+' MPa</h3><p>Résistance moyenne de la sélection.<br>Médiane : '+fmt(s['median'])+' MPa.<br>Écart-type : '+fmt(s['std'])+' MPa.</p><small>Dispersion et niveau se lisent ensemble.</small></article><article class="bs-zone"><span class="bs-step">03 / CE QUE J’EXPLORE</span><h3>Comparer et expliquer</h3><p>Deux groupes distincts, une fiche par essai et les calculs détaillés.</p><small>'+escape(insight_messages(c)[0])+'</small></article></div>')
    with st.container(horizontal=True,gap='medium'):
        for url,label,icon in [('donnees','01 · Examiner les données','database'),('comparateur','05 · Comparer deux groupes','compare_arrows'),('essais','06 · Ouvrir une fiche d’essai','science')]:
            st.page_link(st.session_state.bs_pages[url],label=label,icon=':material/'+icon+':',width='stretch')
    a,b=st.columns([1.3,1])
    with a:
        st.subheader('Le profil de votre sélection')
        chart(histogram(d),'Histogramme des résistances de la sélection')
    with b:
        st.subheader('Ce qui change avec l’âge')
        chart(age_box_chart(d),'Résistances par classes d’âge')
    st.subheader('À retenir avant d’interpréter')
    for i,message in enumerate(insight_messages(c)):
        if i<3:callout('Point de lecture '+str(i+1),message,'purple' if i%2 else 'teal')


def age_box_chart(df):
    return alt.Chart(age_groups(df)).mark_boxplot(color=TEAL,extent=1.5).encode(x=alt.X('Classe d’âge:N',title=None,sort=['1–7 j','8–28 j','29–90 j','> 90 j','Âge inconnu']),y=alt.Y('resistance:Q',title='Résistance (MPa)'))


def compare_groups(frame,key,cutoff):
    if key not in COLS[:-1]+['eau_liant'] or not math.isfinite(float(cutoff)):
        raise ValueError('Variable ou seuil de séparation invalide.')
    known=frame[key].notna();a=frame.loc[known & frame[key].le(cutoff)].copy();b=frame.loc[known & frame[key].gt(cutoff)].copy()
    both=pd.concat([a.assign(groupe='Groupe A'),b.assign(groupe='Groupe B')],ignore_index=True)
    sa,sb=descriptive(a.resistance),descriptive(b.resistance)
    return dict(key=key,cutoff=float(cutoff),a=a,b=b,combined=both,sa=sa,sb=sb,missing=int((~known).sum()),delta=sb['mean']-sa['mean'] if len(a) and len(b) else None)


def comparison_table(comp):
    rows=[]
    for name,s,d in [('Groupe A',comp['sa'],comp['a']),('Groupe B',comp['sb'],comp['b'])]:
        ages=d.age.dropna()
        rows.append({'Groupe':name,'Essais':s['n'],'Moyenne (MPa)':s['mean'],'Médiane (MPa)':s['median'],'Écart-type (MPa)':s['std'],'CV (%)':s['cv'],'Âge moyen (jours)':float(ages.mean()) if len(ages) else None,'Âges renseignés':len(ages)})
    return pd.DataFrame(rows)


def active_comparison(ctx):
    saved=st.session_state.get('bs_comparison')
    if saved and saved['signature']==context_signature(ctx):return compare_groups(ctx['df'],saved['key'],saved['cutoff'])
    return None


def comparison_note(comp):
    cut=fmt(comp['cutoff'],4);variable=LABELS[comp['key']]
    text=f"A : {variable} ≤ {cut} ; B : {variable} > {cut} ({UNITS[comp['key']]}). {comp['missing']} essai(s) non réparti(s), faute de valeur renseignée. "
    if comp['delta'] is None:return text+'Au moins un groupe est vide : la différence de moyennes est non calculable.'
    return text+f"La différence moyenne B − A est de {fmt(comp['delta'])} MPa. Elle est descriptive : les effectifs, les âges et les autres ingrédients peuvent différer. Elle ne mesure pas l’effet causal de cette variable."


def comparison_page():
    heading('05 / COMPARATEUR','Deux groupes, une comparaison transparente.','Séparez les essais sur une variable ; aucun essai ne peut appartenir aux deux groupes.')
    with st.expander('Tester avec quatre essais fictifs'):
        st.markdown('Exemple distinct de vos données : rapports eau/liant **0,40 ; 0,50 ; 0,60 ; 0,70**, résistances **40 ; 35 ; 30 ; 25 MPa**, tous à **28 jours**. Avec un seuil **0,55**, A et B contiennent chacun deux essais : moyennes **37,5** et **27,5 MPa**, différence B − A = **−10 MPa**. Cet exemple construit ne démontre pas une relation générale de formulation.')
        if st.button('Remplacer par les 4 essais fictifs de comparaison',icon=':material/science:',wrap=True):
            fixture=standardize(pd.DataFrame(dict(ciment=[300]*4,laitier=[0]*4,cendres=[0]*4,eau=[120,150,180,210],adjuvant=[0]*4,gravier=[1000]*4,sable=[750]*4,age=[28]*4,resistance=[40,35,30,25])))[0]
            activate(fixture,source_meta('Comparaison guidée · 4 essais fictifs','Exemple fictif','Jeu pédagogique construit pour tester des groupes distincts.'))
            st.rerun()
    ctx=need_data()
    if ctx is None:return
    d=ctx['df'];available=[k for k in COLS[:-1]+['eau_liant'] if d[k].notna().any()]
    if not available:st.info('Renseignez au moins une variable de composition ou l’âge dans Données.');return
    key=st.selectbox('Variable de séparation',available,index=available.index('eau_liant') if 'eau_liant' in available else 0,format_func=lambda k:LABELS[k]+' ('+UNITS[k]+')',key='bs_compare_key',persist_state='session')
    signature=context_signature(ctx)
    cut=st.number_input('Seuil de séparation ('+UNITS[key]+')',min_value=0.,max_value=1e9,value=float(d[key].median()),format='%.4f',key='bs_cut_'+signature[:12]+'_'+key,persist_state='session')
    comp=compare_groups(d,key,cut);sa,sb=comp['sa'],comp['sb']
    st.caption('A contient les valeurs inférieures ou égales au seuil ; B les valeurs strictement supérieures. Les valeurs manquantes restent hors des deux groupes. Les filtres de la barre latérale s’appliquent d’abord.')
    metrics([('Groupe A',fmt(sa['n'],0)+' essais','≤ '+fmt(cut,4)+' '+UNITS[key]),('Groupe B',fmt(sb['n'],0)+' essais','> '+fmt(cut,4)+' '+UNITS[key]),('Écart moyen B − A',fmt(comp['delta'])+' MPa' if comp['delta'] is not None else 'Non calculable','Différence descriptive'),('Non répartis',str(comp['missing']),'Valeur de séparation manquante')])
    grid(comparison_table(comp).round(4),'Effectifs, résistances et âges des deux groupes')
    if len(comp['a']) and len(comp['b']):
        colorscale=alt.Scale(domain=list(GROUP_COLORS),range=list(GROUP_COLORS.values()))
        a,b=st.columns(2)
        with a:
            st.subheader('Les distributions, pas seulement les moyennes')
            chart(alt.Chart(comp['combined']).mark_boxplot(extent=1.5,size=48).encode(x=alt.X('groupe:N',title=None),y=alt.Y('resistance:Q',title='Résistance (MPa)'),color=alt.Color('groupe:N',scale=colorscale,legend=None)),'Dispersion des résistances des groupes A et B')
        with b:
            st.subheader('Une même échelle pour les deux groupes')
            chart(alt.Chart(comp['combined']).mark_bar(opacity=.65).encode(x=alt.X('resistance:Q',bin=alt.Bin(maxbins=15),title='Résistance (MPa)'),y=alt.Y('count():Q',title='Essais'),color=alt.Color('groupe:N',scale=colorscale,title='Groupe'),xOffset='groupe:N',tooltip=['groupe:N',alt.Tooltip('count():Q',title='Essais')]),'Histogramme commun des deux groupes')
    else:st.info('Déplacez le seuil pour obtenir au moins un essai dans chaque groupe. Aucun résultat n’est inventé pour un groupe vide.')
    callout('Comment lire cette différence',comparison_note(comp),'purple')
    if min(sa['n'],sb['n'])<5:st.caption('Au moins un groupe a moins de 5 essais : la comparaison est particulièrement fragile.')
    with st.container(horizontal=True):
        if st.button('Retenir cette comparaison pour le rapport',type='primary',icon=':material/bookmark:',disabled=comp['delta'] is None,width='stretch',wrap=True):
            st.session_state.bs_comparison=dict(signature=signature,key=key,cutoff=cut)
            st.success('Comparaison retenue. Vous pourrez l’inclure dans Rapports. Elle sera invalidée si les données ou les filtres changent.')
        st.download_button('Exporter les groupes en CSV',csv_bytes(comp['combined']),'betonstat-groupes.csv','text/csv',on_click='ignore',icon=':material/download:',width='stretch',wrap=True)


def trial_summary(frame,line):
    selected=frame.loc[frame.ligne.eq(line)]
    if selected.empty:raise ValueError('Cet essai n’appartient pas à la sélection active.')
    row=selected.iloc[0];r=float(row.resistance)
    return dict(row=row,resistance=r,percentile=100*float(frame.resistance.le(r).mean()),n=len(frame),known=int(row[COLS[:7]].notna().sum()),mean=float(frame.resistance.mean()))


def trial_page():
    heading('06 / FICHES D’ESSAIS','Une composition, un résultat, un contexte.','La référence correspond au numéro de ligne du fichier importé, pas à un identifiant de lot inventé.')
    ctx=need_data()
    if ctx is None:return
    d=ctx['df'];lines=d.ligne.astype(int).tolist();sig=context_signature(ctx)[:12];widget='bs_trial_'+sig
    line=st.selectbox('Essai à examiner (ligne source)',lines,format_func=lambda v:'Ligne '+str(v),key=widget)
    pos=lines.index(line)
    def move(target):st.session_state[widget]=target
    with st.container(horizontal=True):
        st.button('Essai précédent',icon=':material/arrow_back:',on_click=move,args=(lines[max(0,pos-1)],),disabled=pos==0,width='stretch')
        st.button('Essai suivant',icon=':material/arrow_forward:',on_click=move,args=(lines[min(len(lines)-1,pos+1)],),disabled=pos==len(lines)-1,width='stretch')
    info=trial_summary(d,line);row=info['row'];age=fmt(row.age,0)+' jours' if pd.notna(row.age) else 'Non renseigné'
    metrics([('Résistance mesurée',fmt(row.resistance)+' MPa','Observation de la ligne '+str(line)),('Âge de l’essai',age,'Sans conversion automatique à 28 jours'),('Position descriptive',fmt(info['percentile'],1)+' %','Part des essais ayant R ≤ cette valeur'),('Composition renseignée',str(info['known'])+' / 7','Constituants connus, sans imputation')])
    comp=pd.DataFrame([{'Constituant':LABELS[k],'Quantité (kg/m³)':row[k],'Couleur':MATERIAL_COLORS[k]} for k in COLS[:7]])
    a,b=st.columns([1.2,1])
    with a:
        st.subheader('Le mélange en un regard')
        valid=comp.dropna(subset=['Quantité (kg/m³)'])
        if len(valid):
            chart(alt.Chart(valid).mark_bar(cornerRadiusEnd=4).encode(y=alt.Y('Constituant:N',sort=comp.Constituant.tolist(),title=None),x=alt.X('Quantité (kg/m³):Q'),color=alt.Color('Constituant:N',scale=alt.Scale(domain=[LABELS[k] for k in COLS[:7]],range=list(MATERIAL_COLORS.values())),legend=None),tooltip=['Constituant:N','Quantité (kg/m³):Q']),'Dosages renseignés des constituants, en kilogrammes par mètre cube')
        else:st.info('Aucun dosage renseigné pour cette ligne. Le résultat de résistance reste consultable.')
    with b:
        st.subheader('Les valeurs de cet essai')
        grid(comp[['Constituant','Quantité (kg/m³)']],'Composition détaillée de l’essai sélectionné')
        callout('Eau / liant massique',fmt(row.eau_liant,4) if pd.notna(row.eau_liant) else 'Non calculable : eau ou constituant du liant manquant. Le vide n’est pas interprété comme zéro.')
    st.subheader('Où se situe cet essai ?')
    marker=alt.Chart(pd.DataFrame({'repere':[float(row.resistance)]})).mark_rule(color=PEACH,strokeWidth=3).encode(x='repere:Q',tooltip=alt.Tooltip('repere:Q',title='Essai sélectionné (MPa)'))
    chart(histogram(d)+marker,'Histogramme de la sélection et résistance de l’essai choisi',250)
    callout('Une position dans ce jeu, sans jugement de qualité',f"{fmt(info['percentile'],1)} % des {info['n']} essais ont une résistance inférieure ou égale à cette valeur, égalités comprises. Ce pourcentage dépend du filtre et du mélange des âges ; il ne prouve ni conformité ni supériorité d’une formulation.",'purple')
    st.download_button('Télécharger cette fiche en CSV',csv_bytes(d.loc[d.ligne.eq(line)]),'betonstat-essai-'+str(line)+'.csv','text/csv',on_click='ignore',icon=':material/download:')


def explain_calculation(indicator,values,unit='MPa'):
    x=pd.Series(values,dtype=float).dropna().to_numpy();s=descriptive(x);n=len(x)
    if not n:st.info('Aucune valeur disponible.');return
    sample=', '.join(fmt(v) for v in x[:10]);st.caption(f"{n} valeur(s) analysée(s) : "+sample+('… (liste abrégée, calcul sur toutes les valeurs)' if n>10 else ''))
    if indicator=='mean':
        st.latex(r'\bar{x}=\frac{\sum x_i}{n}')
        st.write(f"Somme = {fmt(float(x.sum()),4)} {unit} ; n = {n}. Moyenne = {fmt(float(x.sum()),4)} / {n} = {fmt(s['mean'],4)} {unit}.")
        st.write('La moyenne donne le niveau global ; quelques valeurs éloignées peuvent la déplacer.')
    elif indicator=='median':
        ordered=np.sort(x);middle=n//2
        st.write('On classe toutes les valeurs dans l’ordre croissant.')
        if n%2:st.write(f"Effectif impair : la valeur de rang {middle+1} est {fmt(ordered[middle],4)} {unit}.")
        else:st.write(f"Effectif pair : ({fmt(ordered[middle-1],4)} + {fmt(ordered[middle],4)}) / 2 = {fmt(s['median'],4)} {unit}.")
        st.write('La médiane sépare la série ordonnée en deux moitiés, selon la convention usuelle.')
    elif indicator in ['variance','std']:
        st.latex(r's^2=\frac{\sum(x_i-\bar{x})^2}{n-1},\qquad s=\sqrt{s^2}')
        if n<2:st.info('Avec un seul essai, le diviseur n−1 vaut zéro : variance et écart-type d’échantillon ne sont pas calculables.');return
        squares=float(np.square(x-s['mean']).sum())
        st.write(f"Moyenne = {fmt(s['mean'],4)} {unit}. Somme des carrés des écarts = {fmt(squares,4)} {unit}².")
        st.write(f"Variance = {fmt(squares,4)} / {n-1} = {fmt(s['variance'],4)} {unit}². Écart-type = √{fmt(s['variance'],4)} = {fmt(s['std'],4)} {unit}.")
        st.write('L’écart-type décrit la dispersion des résultats autour de la moyenne ; il conserve l’unité des essais.')
    elif indicator=='cv':
        st.latex(r'CV=100\frac{s}{\bar{x}}')
        if s['cv'] is None:st.info('Il faut au moins deux valeurs et une moyenne positive pour ce coefficient.');return
        st.write(f"CV = 100 × {fmt(s['std'],4)} / {fmt(s['mean'],4)} = {fmt(s['cv'],4)} %.")
        st.write('Le coefficient exprime la dispersion relativement à la moyenne. Aucun seuil de qualité universel n’est appliqué.')


def guided_calculations(ctx):
    st.subheader('Comprendre mes calculs')
    st.caption('Ouvrez un indicateur. Le calcul utilise la résistance de la sélection active ou un exemple fictif explicitement choisi.')
    with st.container(horizontal=True,gap='small'):
        for key,title in [('mean','Moyenne'),('median','Médiane'),('variance','Variance'),('std','Écart-type'),('cv','CV')]:
            with st.popover('Comprendre : '+title,icon=':material/help:',width='stretch',wrap=True):
                mode=st.selectbox('Valeurs utilisées pour '+title,['Sélection active','Exemple fictif : 20, 25, 30, 35, 40'],key='bs_explain_'+key)
                if mode.startswith('Exemple'):st.caption('Exemple fictif de calcul, distinct de votre jeu actif.')
                explain_calculation(key,ctx['df'].resistance if mode=='Sélection active' else [20,25,30,35,40])


def default_report_options(ctx):
    return dict(title='Analyse statistique des résistances du béton',student='',institution='',date=datetime.now(timezone.utc).strftime('%Y-%m-%d'),graphs=['histogramme','ages'],relation_key='ciment',comparison=None,conclusion=auto_conclusion(ctx),assumptions='Statistiques descriptives sur la sélection active. Valeurs manquantes non imputées. Comparabilité et indépendance des essais à vérifier ; absence d’identifiants de lots dans le jeu UCI. Aucun critère réglementaire intégré.')


def report_signature(ctx,options):
    return hashlib.sha256((context_signature(ctx)+json.dumps(options,ensure_ascii=False,sort_keys=True)).encode()).hexdigest()


def reports_page_v2():
    heading('07 / RAPPORTS','Des résultats prêts à expliquer.','Personnalisez la présentation, choisissez les graphiques et relisez la conclusion avant de préparer vos documents.')
    ctx=current()
    if not ctx['raw'].empty:
        with st.container(horizontal=True):
            st.download_button('Sauvegarde complète JSON',backup_bytes(ctx['raw'],ctx['meta'],ctx['rules'],ctx['filters']),'betonstat-sauvegarde.json','application/json',on_click='ignore',icon=':material/save:',width='stretch',wrap=True)
            st.download_button('Sélection CSV',csv_bytes(ctx['df']),'betonstat-selection.csv','text/csv',on_click='ignore',icon=':material/download:',width='stretch',wrap=True)
        defaults=default_report_options(ctx);sig=context_signature(ctx);comp=active_comparison(ctx)
        st.subheader('Présentation du document')
        title=st.text_input('Titre du rapport',value=defaults['title'],max_chars=180,key='bs_report_title',persist_state='session')
        a,b=st.columns(2)
        student=a.text_input('Nom de l’étudiant ou du groupe (facultatif)',max_chars=120,key='bs_report_student',persist_state='session')
        institution=b.text_input('Établissement / filière (facultatif)',max_chars=160,key='bs_report_institution',persist_state='session')
        report_date=st.date_input('Date du rapport',key='bs_report_date',persist_state='session')
        st.subheader('Graphiques à inclure')
        selected=st.multiselect('Choix des graphiques',['histogramme','ages','relation'],default=['histogramme','ages'],format_func=lambda v:dict(histogramme='Histogramme des résistances',ages='Résistance selon les classes d’âge',relation='Nuage de points et droite descriptive')[v],key='bs_report_graphs',persist_state='session')
        available=[k for k in COLS[:-1]+['eau_liant'] if ctx['df'][k].notna().sum()>=2]
        relation_key='ciment'
        if 'relation' in selected:
            if available:relation_key=st.selectbox('Variable du nuage de points',available,format_func=lambda k:LABELS[k],key='bs_report_relation',persist_state='session')
            else:st.info('Pas assez de données de composition ou d’âge pour un nuage de points ; le PDF expliquera cette absence.')
        use_comparison=st.checkbox('Inclure la comparaison retenue',disabled=comp is None,key='bs_report_compare_'+sig[:12],help='Retenez d’abord une comparaison sur la page Comparateur.',persist_state='session')
        if comp:st.caption(comparison_note(comp))
        else:st.caption('Aucune comparaison retenue pour cette sélection. Un changement de données ou de filtres invalide une comparaison précédente.')
        st.subheader('Votre interprétation')
        conclusion_key='bs_report_conclusion_'+sig[:12]
        if st.button('Régénérer la proposition de conclusion',icon=':material/refresh:'):
            st.session_state[conclusion_key]=defaults['conclusion']
        conclusion=st.text_area('Conclusion modifiable',value=defaults['conclusion'],max_chars=5000,height=170,key=conclusion_key,persist_state='session')
        assumptions=st.text_area('Hypothèses et limites à mentionner',value=defaults['assumptions'],max_chars=3000,height=110,key='bs_report_assumptions',persist_state='session')
        options=dict(title=title.strip() or defaults['title'],student=student,institution=institution,date=report_date.isoformat(),graphs=selected,relation_key=relation_key,comparison=dict(key=comp['key'],cutoff=comp['cutoff']) if comp and use_comparison else None,conclusion=conclusion,assumptions=assumptions)
        rsig=report_signature(ctx,options)
        if st.button('Préparer le rapport PDF et Excel',type='primary',icon=':material/picture_as_pdf:'):
            with st.spinner('Mise en page du rapport et des graphiques…'):
                st.session_state.bs_reports=(rsig,pdf_bytes(ctx,options),xlsx_bytes(ctx))
        saved=st.session_state.get('bs_reports')
        if saved and saved[0]==rsig:
            with st.container(horizontal=True):
                st.download_button('Télécharger le rapport PDF',saved[1],'betonstat-rapport.pdf','application/pdf',on_click='ignore',width='stretch',wrap=True)
                st.download_button('Télécharger le classeur Excel',saved[2],'betonstat-analyse.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',on_click='ignore',width='stretch',wrap=True)
        elif saved:st.info('Les réglages ont changé. Préparez à nouveau les documents pour télécharger une version à jour.')
        st.caption('Le PDF contient la page de présentation, les résultats, les graphiques choisis, la comparaison éventuellement retenue et votre conclusion. Le classeur contient les cinq tables de données et de traçabilité ; il ne reprend pas la mise en page du PDF. Sauvegardez le JSON pour conserver les données ; le texte du rapport reste dans la session.')
    st.subheader('Restaurer une sauvegarde')
    up=st.file_uploader('Sauvegarde BétonStat JSON (12 Mo maximum)',type='json',max_upload_size=12,key='bs_restore_file')
    if up:
        try:
            raw,meta,rules,filters=read_backup(up.getvalue());st.write(f"{len(raw)} lignes · {meta['name']} · {meta['kind']}")
            ok=st.checkbox('Remplacer le jeu actif et ses réglages par cette sauvegarde',key='bs_restore_confirm')
            if st.button('Restaurer la sauvegarde',disabled=not ok,type='primary'):
                activate(raw,meta,rules,filters);st.rerun()
        except (ValueError,TypeError,KeyError) as exc:st.error(str(exc))
    callout('Rapport et sauvegarde ont deux rôles','Le PDF présente votre analyse. Le JSON restaure les données brutes, leur provenance et les filtres ; il ne contient pas les textes de présentation du rapport.')


def pdf_graph(ctx,kind,key='ciment',comparison=None):
    from reportlab.graphics.shapes import Drawing,Rect,Line,String,Circle
    w,h=515,245;draw=Drawing(w,h);left,right,bottom,top=55,495,48,216
    draw.add(Rect(0,0,w,h,fillColor=colors.HexColor('#F6F5FA'),strokeColor=None,rx=10,ry=10))
    def text(x,y,value,size=8,color=INK,anchor='start'):
        draw.add(String(x,y,str(value),fontName='BSReport',fontSize=size,fillColor=colors.HexColor(color),textAnchor=anchor))
    def axes(xmin,xmax,ymin,ymax,xlabel,ylabel):
        if xmax<=xmin:xmin-=.5;xmax+=.5
        if ymax<=ymin:ymin-=.5;ymax+=.5
        X=lambda v:left+(float(v)-xmin)/(xmax-xmin)*(right-left)
        Y=lambda v:bottom+(float(v)-ymin)/(ymax-ymin)*(top-bottom)
        for v in np.linspace(ymin,ymax,5):
            y=Y(v);draw.add(Line(left,y,right,y,strokeColor=colors.HexColor('#DDD8E9'),strokeWidth=.5));text(left-7,y-3,fmt(v,1),anchor='end')
        for v in np.linspace(xmin,xmax,5):text(X(v),bottom-16,fmt(v,2),anchor='middle')
        draw.add(Line(left,bottom,right,bottom,strokeColor=colors.HexColor(INK),strokeWidth=.7))
        text((left+right)/2,10,xlabel,9,anchor='middle');text(left,231,ylabel,9)
        return X,Y
    df=ctx['df']
    if df.empty:text(30,125,'Aucun essai analysable dans cette sélection.',11);return draw
    if kind=='histogramme':
        counts,edges=np.histogram(df.resistance.to_numpy(),bins=min(16,max(3,int(np.sqrt(len(df))))))
        X,Y=axes(edges[0],edges[-1],0,max(int(counts.max())*1.15,1),'Résistance (MPa)','Nombre d’essais')
        for n,l,r in zip(counts,edges[:-1],edges[1:]):draw.add(Rect(X(l)+1,Y(0),max(X(r)-X(l)-2,.5),Y(n)-Y(0),fillColor=colors.HexColor(PURPLE),strokeColor=None))
    elif kind=='ages':
        groups=age_groups(df);labels=['1–7 j','8–28 j','29–90 j','> 90 j','Âge inconnu'];groups=[(name,g.resistance.to_numpy()) for name in labels if len(g:=groups.loc[groups['Classe d’âge'].eq(name)])]
        X,Y=axes(0,len(groups),0,max(float(df.resistance.max())*1.15,1),'Classes d’âge (jours)','Résistance (MPa) : boîte = Q1 à Q3 ; trait = médiane')
        # Masque les graduations numériques de l’axe des catégories, puis libelle chaque boîte.
        draw.add(Rect(20,21,490,19,fillColor=colors.HexColor('#F6F5FA'),strokeColor=None))
        for i,(name,x) in enumerate(groups):
            x0=X(i+.5);q1,med,q3=np.quantile(x,[.25,.5,.75]);iqr=q3-q1;inside=x[(x>=q1-1.5*iqr)&(x<=q3+1.5*iqr)];low,high=float(inside.min()),float(inside.max())
            draw.add(Line(x0,Y(low),x0,Y(high),strokeColor=colors.HexColor(TEAL)))
            for whisker in [low,high]:draw.add(Line(x0-10,Y(whisker),x0+10,Y(whisker),strokeColor=colors.HexColor(TEAL)))
            for extreme in x[(x<low)|(x>high)]:draw.add(Circle(x0,Y(extreme),1.7,fillColor=None,strokeColor=colors.HexColor(TEAL),strokeWidth=.6))
            draw.add(Rect(x0-21,Y(q1),42,max(Y(q3)-Y(q1),.7),fillColor=colors.HexColor('#C5E2DF'),strokeColor=colors.HexColor(TEAL)))
            draw.add(Line(x0-21,Y(med),x0+21,Y(med),strokeColor=colors.HexColor(TEAL),strokeWidth=2));text(x0,bottom-16,name,8,anchor='middle')
    elif kind=='relation':
        data=paired(df,key);r=relationship(df,key)
        if data.empty:text(30,125,'Aucune paire complète pour ce graphique.',11);return draw
        X,Y=axes(float(data[key].min()),float(data[key].max()),float(data.resistance.min()),float(data.resistance.max()),LABELS[key]+' ('+UNITS[key]+')','Résistance (MPa)')
        indices=np.linspace(0,len(data)-1,min(1000,len(data)),dtype=int)
        for _,row in data.iloc[indices].iterrows():draw.add(Circle(X(row[key]),Y(row.resistance),1.7,fillColor=colors.HexColor(TEAL),strokeColor=None))
        if r['slope'] is not None:
            # Limite la droite à la fenêtre du graphique, même si elle dépasse les observations.
            xmin,xmax=float(data[key].min()),float(data[key].max());xs=np.linspace(xmin,xmax,300);ys=r['intercept']+r['slope']*xs;inside=(ys>=data.resistance.min())&(ys<=data.resistance.max())
            if inside.sum()>1:draw.add(Line(X(xs[inside][0]),Y(ys[inside][0]),X(xs[inside][-1]),Y(ys[inside][-1]),strokeColor=colors.HexColor(PEACH),strokeWidth=1.6))
    elif kind=='comparaison' and comparison:
        comp=compare_groups(df,comparison['key'],comparison['cutoff']);means=[comp['sa']['mean'],comp['sb']['mean']]
        valid=[v for v in means if v is not None];X,Y=axes(0,2,0,max(valid+[1])*1.25,'Groupes distincts','Résistance moyenne (MPa)')
        draw.add(Rect(20,21,490,19,fillColor=colors.HexColor('#F6F5FA'),strokeColor=None))
        for i,(name,value) in enumerate(zip(GROUP_COLORS,means)):
            if value is not None:
                x=X(i+.5);draw.add(Rect(x-45,Y(0),90,Y(value)-Y(0),fillColor=colors.HexColor(GROUP_COLORS[name]),strokeColor=None));text(x,Y(value)+9,fmt(value)+' MPa',9,anchor='middle');text(x,bottom-16,name,9,anchor='middle')
    return draw


def report_additions(items,ctx,options,styles,para,tb):
    from reportlab.platypus import KeepTogether,PageBreak
    graphs=options.get('graphs',[])
    if graphs or options.get('comparison'):items.append(PageBreak());para('Graphiques et comparaisons','Heading2')
    labels={'histogramme':('Distribution des résistances','Comptage observé ; les classes sont communes à toute la sélection.'),'ages':('Résistance et âge','Boîtes : quartiles ; trait : médiane ; moustaches : valeurs extrêmes à moins de 1,5 IQR des quartiles. Les points au-delà restent représentés, sans suppression.'),'relation':('Relation exploratoire avec '+LABELS.get(options['relation_key'],'Ciment'),'Nuage de points : au plus 1 000 observations régulièrement sélectionnées dans l’ordre des lignes. La droite est ajustée sur toutes les paires complètes. Aucune causalité ni validation prédictive.')}
    for kind in graphs:
        if kind not in labels:continue
        title,caption=labels[kind]
        items.append(KeepTogether([Paragraph(escape(title),styles['Heading2']),pdf_graph(ctx,kind,options['relation_key']),Spacer(1,5),Paragraph(escape(caption),styles['BodyText']),Spacer(1,14)]))
    if options.get('comparison'):
        items.append(PageBreak())
        c=compare_groups(ctx['df'],options['comparison']['key'],options['comparison']['cutoff'])
        items.append(KeepTogether([Paragraph('Comparaison retenue',styles['Heading2']),pdf_graph(ctx,'comparaison',comparison=options['comparison']),Spacer(1,8)]))
        para(comparison_note(c));tb(['Groupe','n','Moyenne (MPa)','Écart-type (MPa)','Âge moyen (j)'],[[r['Groupe'],str(r['Essais']),fmt(r['Moyenne (MPa)']),fmt(r['Écart-type (MPa)']),fmt(r['Âge moyen (jours)'])] for r in comparison_table(c).to_dict('records')],[100,45,120,125,95])
    items.append(PageBreak());para('Interprétation et hypothèses','Heading2')
    para('Conclusion rédigée ou validée par l’utilisateur','Heading2')
    for line in (options.get('conclusion') or 'Aucune conclusion renseignée.').splitlines():
        if line.strip():para(line)
    para('Hypothèses et limites déclarées','Heading2')
    for line in (options.get('assumptions') or 'Aucune hypothèse supplémentaire renseignée.').splitlines():
        if line.strip():para(line)
    para('Repères automatiques à relire','Heading2')
    for message in insight_messages(ctx):para(message)
    para('Le texte libre appartient à l’utilisateur. Les observations, modèles et interprétations doivent rester distingués. Ce document ne certifie ni conformité réglementaire ni sécurité d’un ouvrage.')


def apply_style():
    st.html('''<style>
    :root{--bs-ink:#21183A;--bs-teal:#006D77;--bs-purple:#6651D5}
    [data-testid="stAppViewContainer"]{background:radial-gradient(ellipse at 94% 0%,#E8E2FB 0,transparent 40%),#F5F6FA;color:#21183A;font-family:'Segoe UI',Arial,sans-serif}
    [data-testid="stAppViewContainer"] p,[data-testid="stAppViewContainer"] h1,[data-testid="stAppViewContainer"] h2,[data-testid="stAppViewContainer"] h3,[data-testid="stAppViewContainer"] input,[data-testid="stAppViewContainer"] button{font-family:'Segoe UI',Arial,sans-serif}
    [data-testid="stMainBlockContainer"]{max-width:1460px;padding-top:4.5rem;padding-bottom:3rem}
    [data-testid="stHeader"]{background:rgba(245,246,250,.96)}
    [data-testid="stSidebar"]{background:#21183A;color:#F4F1FF;border-right:1px solid #44345F}
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,[data-testid="stSidebar"] label,[data-testid="stSidebar"] h3{color:#F4F1FF}
    [data-testid="stSidebarNavLink"] span{color:#F4F1FF!important}
    [data-testid="stSidebarNavLink"][aria-current="page"]{background:#49396C;border-left:3px solid #D3EC9A}
    [data-testid="stSidebar"] [data-baseweb="select"]>div{background:#392C53;color:#F4F1FF}
    [data-testid="stSidebar"] input{color:#F4F1FF!important}
    [data-testid="stSidebar"] [data-baseweb="select"] *{color:#F4F1FF!important}
    [data-testid="stSidebar"] [data-testid="stExpander"]{background:#392C53;border-color:#5B4C75}
    [data-testid="stCaptionContainer"]{opacity:1!important}
    [data-testid="stCaptionContainer"] p{color:#585269}
    [data-testid="stSidebar"] [data-testid="stCaptionContainer"] p{color:#D2C9E6}
    [data-testid="stWidgetLabel"] p{white-space:normal!important;overflow:visible!important}
    .stButton button,.stDownloadButton button{border-radius:12px;min-height:44px;border:1px solid #C8C2DA;font-weight:650}
    button[kind="primary"]{background:#006D77!important;color:white!important;border-color:#006D77!important}
    .stButton button:hover,.stDownloadButton button:hover{border-color:#6651D5;box-shadow:0 3px 12px #6651D51C}
    [data-testid="stForm"],[data-testid="stExpander"]{background:#FFFFFFCC;border:1px solid #DED9EA;border-radius:15px}
    h1,h2,h3{color:#21183A;letter-spacing:-.03em}
    .bs-brand{padding:14px 0 20px}.bs-brand strong{font-size:29px;letter-spacing:-1px;color:white}.bs-brand small{display:block;color:#D3EC9A;letter-spacing:2px;font-size:10px;margin-top:5px}
    .bs-mast{display:flex;justify-content:space-between;align-items:center;gap:15px;margin-bottom:18px}.bs-mast span{font-size:11px;letter-spacing:2px;color:#635970}.bs-chip{background:#E5F1EF;color:#00545C;padding:8px 12px;border-radius:24px;font-size:12px;max-width:100%;overflow-wrap:anywhere}
    .bs-hero{display:flex;align-items:center;justify-content:space-between;gap:20px;padding:30px 34px;background:#21183A;color:white;border-radius:24px;position:relative;overflow:hidden;margin-bottom:20px;border-bottom:5px solid #D3EC9A}
    .bs-hero h1{color:white;font-size:42px;line-height:1.12;margin:12px 0 14px;padding:0}.bs-hero p{font-size:15px;line-height:1.6;color:#DAD3EA;margin:0}.bs-eyebrow{color:#D3EC9A;font-size:10px;letter-spacing:2px;font-weight:650}
    .bs-orbit{flex:0 0 180px;position:relative;width:180px;height:180px;border:1px solid #6651D5;border-radius:50%;background:radial-gradient(circle,#6651D560,transparent 70%);transform:rotate(-12deg)}
    .bs-orbit:after{content:'';position:absolute;inset:20px;border:1px dashed #9B8BCD;border-radius:50%}.bs-orbit span{position:absolute;left:54px;top:33px;font-family:Georgia,serif;font-size:80px;color:#D3EC9A}.bs-orbit i{position:absolute;right:-3px;bottom:28px;font-size:16px;background:#006D77;padding:9px;border-radius:8px;font-style:normal}.bs-orbit b{position:absolute;top:2px;right:11px;font-size:32px;color:#EFBF9F}.bs-orbit em{position:absolute;left:-9px;bottom:7px;font-size:43px;color:#BEB0F0;font-family:Georgia,serif}
    .bs-metrics{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:13px;margin:15px 0 22px}.bs-metric{background:white;border:1px solid #DDD8E9;border-top:4px solid #6651D5;border-radius:15px;padding:16px 18px;min-width:0}.bs-metric:nth-child(2){border-top-color:#006D77}.bs-metric:nth-child(3){border-top-color:#D56A3D}.bs-metric:nth-child(4){border-top-color:#93AF4E}.bs-metric span{display:block;font-size:12px;color:#585269}.bs-metric strong{display:block;font-size:25px;line-height:1.35;margin:8px 0;color:#21183A;overflow-wrap:anywhere}.bs-metric small{display:block;color:#585269;font-size:11px;line-height:1.5}
    .bs-note{background:#E6F2F0;border-left:4px solid #006D77;padding:16px 20px;border-radius:0 13px 13px 0;margin:16px 0}.bs-note strong{color:#004E56}.bs-note p{font-size:14px;color:#294F52;margin:7px 0 0;line-height:1.6}.bs-note.purple{background:#EEEAFA;border-color:#6651D5}.bs-note.purple strong,.bs-note.purple p{color:#473B70}
    .bs-heading span{font-size:11px;letter-spacing:2px;color:#006D77}.bs-heading h1{font-size:33px;line-height:1.2;margin:10px 0;padding:0}.bs-heading p{font-size:15px;color:#585269;margin-bottom:22px}
    .bs-path{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin:22px 0}.bs-path>div{padding:20px;background:#FFFFFF;border:1px solid #DED9EA;border-radius:15px}.bs-path b{color:#6651D5;font-size:24px}.bs-path strong{display:block;margin:9px 0;font-size:16px;color:#21183A}.bs-path span{font-size:12px;color:#585269;line-height:1.6}
    @media(max-width:900px){.bs-metrics{grid-template-columns:repeat(2,minmax(0,1fr))}.bs-path{grid-template-columns:repeat(2,1fr)}.bs-hero h1{font-size:35px}.bs-orbit{flex-basis:145px;width:145px;height:145px}.bs-orbit span{left:40px;top:22px;font-size:65px}}
    @media(max-width:640px){[data-testid="stMainBlockContainer"]{padding:4.5rem 1rem 2rem}.bs-hero{padding:24px 20px;border-radius:17px}.bs-hero h1{font-size:31px}.bs-orbit{display:none}.bs-heading h1{font-size:28px}.bs-metric{padding:13px 11px}.bs-metric strong{font-size:21px}.bs-mast{align-items:flex-start;flex-direction:column;gap:8px}.bs-path>div{padding:15px}}
    @media(max-width:360px){.bs-metrics,.bs-path{grid-template-columns:1fr}.bs-hero h1{font-size:28px}}

    .bs-hero-live{background:linear-gradient(115deg,#21183A 8%,#3F2B69 70%,#006D77 155%);padding:32px 36px}
    .bs-hero-stat{flex:0 0 195px;padding:18px 22px;border:1px solid #9380BE80;background:#FFFFFF0A;border-radius:18px;color:#D3EC9A}
    .bs-hero-stat small{display:block;font-size:10px;letter-spacing:2px;color:#E5DEF4}.bs-hero-stat strong{display:block;font-size:48px;line-height:1.2;font-weight:700;color:#D3EC9A;margin:7px 0}.bs-hero-stat span{font-size:12px;color:#E5DEF4}.bs-spark{display:flex;align-items:flex-end;gap:4px;width:100%;height:42px;margin-top:8px}.bs-spark span{display:block;flex:1;min-width:0;background:#D3EC9A;border-radius:2px 2px 0 0}
    .bs-zones{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:16px;margin:4px 0 22px}
    .bs-zone{position:relative;padding:22px;border-radius:18px;background:#E5F1EF;border:1px solid #C6E1DD;border-top:4px solid #006D77;min-width:0}
    .bs-zone:nth-child(2){background:#EDE8FA;border-color:#D8CEF1;border-top-color:#6651D5}.bs-zone:nth-child(3){background:#F9EEE7;border-color:#EED4C2;border-top-color:#B15C35}
    .bs-step{font-size:10px;letter-spacing:1.4px;font-weight:700;color:#285B5B}.bs-zone:nth-child(2) .bs-step{color:#534081}.bs-zone:nth-child(3) .bs-step{color:#85502F}
    .bs-zone h3{font-size:22px;line-height:1.25;margin:13px 0}.bs-zone p{font-size:14px;color:#3B4055;line-height:1.7;margin:10px 0}.bs-zone small{font-size:12px;color:#4F4D5C;line-height:1.6;display:block}.bs-zone b{font-weight:600}
    [data-testid="stPageLink"]{padding:10px;background:#FFFFFF;border:1px solid #D8D1E7;border-radius:12px;min-height:48px}
    [data-testid="stSidebarNavLink"]{min-height:38px;border-radius:8px}
    [data-testid="stPopover"] button{border-radius:11px;min-height:44px;border-color:#CEC6E1;color:#4D3B77;background:#F4F0FD}
    @media(max-width:1000px){.bs-zones{grid-template-columns:1fr}.bs-zone{padding:18px 22px}.bs-zone h3{margin:8px 0}.bs-zone p{margin:7px 0}.bs-hero-stat{flex-basis:170px}}
    @media(max-width:640px){.bs-hero-live{padding:24px 20px;align-items:flex-start;flex-direction:column}.bs-hero-stat{width:100%;flex-basis:auto;box-sizing:border-box;display:grid;grid-template-columns:1fr 1fr;align-items:center;gap:4px 15px;padding:15px 18px}.bs-hero-stat strong{font-size:38px;grid-row:2;grid-column:1}.bs-hero-stat span{grid-column:1;grid-row:3}.bs-hero-stat .bs-spark{grid-column:2;grid-row:1/4;height:60px;margin:0}.bs-hero-stat small{letter-spacing:1px;font-size:9px}.bs-zones{gap:12px}.bs-zone h3{font-size:21px}}
    </style>''')


def main():
    st.set_page_config(page_title=APP_NAME+' · Atelier statistique',page_icon=':material/science:',layout='wide',initial_sidebar_state='auto')
    for key,value in [('bs_raw',pd.DataFrame(columns=COLS)),('bs_meta',source_meta('Aucune donnée','Aucune source')),('bs_revision',0),('bs_defaults',dict(rules=dict(dedup=False,complete=False),filters=dict(age_mode='Tous âges',age_low=1.,age_high=365.,threshold=30.)))]:
        if key not in st.session_state:st.session_state[key]=value
    apply_style()
    pages=[st.Page(overview_v2,title='00 · Vue d’ensemble',url_path='accueil',default=True,icon=':material/dashboard:'),
        st.Page(data_page,title='01 · Données',url_path='donnees',icon=':material/database:'),
        st.Page(statistics_page,title='02 · Statistiques',url_path='statistiques',icon=':material/functions:'),
        st.Page(distributions_page,title='03 · Distributions',url_path='distributions',icon=':material/stacked_bar_chart:'),
        st.Page(relations_page,title='04 · Relations',url_path='relations',icon=':material/scatter_plot:'),
        st.Page(comparison_page,title='05 · Comparateur',url_path='comparateur',icon=':material/compare_arrows:'),
        st.Page(trial_page,title='06 · Fiches d’essais',url_path='essais',icon=':material/science:'),
        st.Page(reports_page_v2,title='07 · Rapports',url_path='rapports',icon=':material/description:'),
        st.Page(methods_page,title='08 · Méthodes et tests',url_path='methodes',icon=':material/school:')]
    st.session_state.bs_pages={p.url_path:p for p in pages}
    with st.sidebar:st.html('<div class="bs-brand"><strong>BétonStat<span style="color:#D3EC9A">.</span></strong><small>LE LABORATOIRE DES DONNÉES</small></div>')
    page=st.navigation(pages,position='sidebar',expanded=True)
    rev=st.session_state.bs_revision;defaults=st.session_state.bs_defaults
    with st.sidebar:
        st.subheader('Sélection d’analyse')
        age_mode=st.selectbox('Âge des essais',['Tous âges','28 jours','Plage d’âge'],index=['Tous âges','28 jours','Plage d’âge'].index(defaults['filters']['age_mode']),key='bs_age_mode_'+str(rev))
        lo,hi=defaults['filters']['age_low'],defaults['filters']['age_high']
        if age_mode=='Plage d’âge':
            lo=st.number_input('Âge minimum (jours)',0.,1e9,float(lo),1.,key='bs_age_low_'+str(rev));hi=st.number_input('Âge maximum (jours)',0.,1e9,float(hi),1.,key='bs_age_high_'+str(rev))
        threshold=st.number_input('Seuil exploratoire (MPa)',0.,1000.,float(defaults['filters']['threshold']),1.,key='bs_threshold_'+str(rev),help='Comparaison statistique R < seuil, sans déclaration de conformité.')
        with st.expander('Règles de nettoyage'):
            dedup=st.checkbox('Retirer les doublons complets',value=defaults['rules']['dedup'],key='bs_dedup_'+str(rev),wrap=True)
            complete=st.checkbox('Exiger composition et âge complets',value=defaults['rules']['complete'],key='bs_complete_'+str(rev),wrap=True)
        st.caption('Les valeurs invalides sont exclues et expliquées. Aucune imputation automatique, aucune suppression automatique des valeurs extrêmes.')
        st.caption('Version '+VERSION+' · Outil pédagogique')
    raw=st.session_state.bs_raw;rules=dict(dedup=dedup,complete=complete);filters=dict(age_mode=age_mode,age_low=lo,age_high=hi,threshold=threshold)
    if raw.empty:
        cleaned=pd.DataFrame(columns=['ligne']+COLS+['eau_liant']);audit=pd.DataFrame();quality=dict(raw=0,valid=0,excluded=0,duplicates=0)
    else:cleaned,audit,quality=clean_data(raw,dedup,complete)
    df=cleaned.copy()
    if age_mode=='28 jours':df=df.loc[df['age'].eq(28)]
    elif age_mode=='Plage d’âge':
        if lo>hi:st.error('La borne d’âge minimum dépasse la borne maximum. La sélection est vide.')
        df=df.loc[df['age'].between(lo,hi)]
    ctx=dict(raw=raw,clean=cleaned,df=df,audit=audit,quality=quality,meta=st.session_state.bs_meta,rules=rules,filters=filters)
    st.session_state.bs_context=ctx
    st.html(f'<div class="bs-mast"><span>BÉTONSTAT / ATELIER STATISTIQUE</span><div class="bs-chip">{escape(ctx["meta"]["kind"])} · {len(df)} essai(s) sélectionné(s)</div></div>')
    if not raw.empty:
        st.caption(ctx['meta']['name']+' · '+age_mode+' · Les graphiques et tableaux suivent cette sélection.')
        if age_mode=='Tous âges' and df['age'].nunique()>1:st.caption('Âges mélangés : moyenne globale descriptive, pas une résistance de référence à 28 jours.')
    page.run()


if __name__=='__main__':
    main()
