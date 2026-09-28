import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Polygon, Circle, FancyArrowPatch
import numpy as np
import os
OUT=os.path.join(os.path.dirname(os.path.abspath(__file__)),"figures")+"/"
plt.rcParams.update({"font.family":"DejaVu Sans","font.size":9})
INK="#1f2937"; POW="#b91c1c"; SIG="#475569"; DONE="#15803d"; NOW="#b45309"; TODO="#6b7280"

def box(ax,x,y,w,h,text,fc="white",ec=INK,lw=1.2,fs=8.5,bold=False,tc=INK):
    ax.add_patch(FancyBboxPatch((x,y),w,h,boxstyle="round,pad=0.02,rounding_size=0.08",fc=fc,ec=ec,lw=lw))
    ax.text(x+w/2,y+h/2,text,ha="center",va="center",fontsize=fs,color=tc,fontweight="bold" if bold else "normal",linespacing=1.3)
def arrow(ax,x1,y1,x2,y2,c=SIG,lw=1.2,style="-|>",ls="-"):
    ax.add_patch(FancyArrowPatch((x1,y1),(x2,y2),arrowstyle=style,mutation_scale=10,color=c,lw=lw,linestyle=ls,shrinkA=0,shrinkB=0))
def canvas(w,h,W,H):
    fig,ax=plt.subplots(figsize=(w,h)); ax.set_xlim(0,W); ax.set_ylim(0,H); ax.axis("off"); return fig,ax
def save(fig,name):
    fig.savefig(OUT+name,dpi=220,bbox_inches="tight",facecolor="white"); plt.close(fig)

# Fig: overall block diagram
fig,ax=canvas(7.2,6.0,12,10)
box(ax,0.2,8.6,2.2,1.0,"3S LiPo\n11.1 V 2200 mAh",fs=7.8,fc="#fef2f2",ec=POW)
box(ax,2.9,8.6,1.8,1.0,"20 A blade\nfuse",fc="#fef2f2",ec=POW)
box(ax,5.2,8.6,1.8,1.0,"20 A rocker\nswitch",fc="#fef2f2",ec=POW)
box(ax,7.5,8.6,2.4,1.0,"Split point /\nstar ground",fc="#fef2f2",ec=POW)
for a,b in [(2.4,2.9),(4.7,5.2),(7.0,7.5)]: arrow(ax,a,9.1,b,9.1,c=POW,lw=1.8)
conv=[("Conv. A  XL4016\n6.0 V (servo rail A)",0.2),("Conv. B  XL4016\n6.0 V (servo rail B)",3.2),("Conv. C  LM2596\n5.0 V (logic)",6.2),("Conv. D  LM2596\n5.0 V (camera only)",9.2)]
for t,x in conv:
    box(ax,x,6.9,2.6,1.0,t,fc="#fef2f2",ec=POW,fs=8)
    ax.plot([8.7,8.7],[8.6,8.25],color=POW,lw=1.8); ax.plot([1.5,10.5],[8.25,8.25],color=POW,lw=1.8)
    arrow(ax,x+1.3,8.25,x+1.3,7.9,c=POW,lw=1.8)
ax.text(6,7.95+0.45,"converter tray (under body plate)",ha="center",fontsize=7.5,color=POW,style="italic")
box(ax,0.2,3.6,8.6,2.6,"",fc="#f8fafc",ec=INK,lw=1.6)
ax.text(4.5,5.95,"CUSTOM PCB  —  Arduino Uno shield, approx. 114 × 84 mm",ha="center",fontsize=8.5,fontweight="bold",color=INK)
box(ax,0.45,4.7,2.5,0.95,"6 V zone A\nright legs + camera",fs=6.8,ec=POW)
box(ax,3.2,4.7,2.5,0.95,"6 V zone B\nleft legs",fs=7,ec=POW)
box(ax,5.95,4.7,2.6,0.95,"+5 V logic zone\nUno, drivers, sensors",fs=7,ec=POW)
box(ax,0.45,3.8,3.9,0.7,"2 × PCA9685 + 20 servo sockets",fs=7)
box(ax,4.55,3.8,4.0,0.7,"Sensor headers, LEDs, buzzer,\nbattery sense",fs=6.8)
for x,xe in [(1.5,1.7),(4.5,4.45),(7.5,7.25)]: arrow(ax,xe,6.9,xe,5.65,c=POW,lw=1.6)
box(ax,9.2,4.3,2.6,1.3,"ESP32-CAM\n(own 5 V node)\nWi-Fi video",fc="#eff6ff",ec="#1d4ed8")
arrow(ax,10.5,6.9,10.5,5.6,c=POW,lw=1.6)
ax.annotate("",xy=(8.8,4.95),xytext=(9.2,4.95),arrowprops=dict(arrowstyle="<|-|>",color=SIG,lw=1.2))
ax.text(9.0,5.15,"serial\n9600",ha="center",fontsize=6.5,color=SIG)
box(ax,2.5,2.25,4.2,0.95,"ARDUINO UNO R3\nmain controller, under shield",fc="#f1f5f9",fs=7.2,bold=True)
ax.plot([4.6,4.6],[3.2,3.6],color=SIG,lw=3)
box(ax,0.0,0.4,2.6,1.3,"18 × MG996R\n6 legs × 3 joints",fs=8)
box(ax,2.9,0.4,2.2,1.3,"2 × SG90\ncamera pan / tilt",fs=8)
box(ax,5.4,0.4,2.2,1.3,"3 × HC-SR04\nfront, left, right",fs=8)
box(ax,7.9,0.4,1.9,1.3,"MPU6050\nIMU",fs=8)
box(ax,10.1,0.4,1.8,1.3,"Phone /\nbrowser",fc="#eff6ff",ec="#1d4ed8",fs=8)
arrow(ax,1.3,3.6,1.3,1.7); arrow(ax,3.9,3.6,4.0,1.7); arrow(ax,6.5,3.6,6.5,1.7,style="<|-|>"); arrow(ax,8.1,3.6,8.8,1.7,style="<|-|>")
arrow(ax,10.9,4.3,11.0,1.7,c="#1d4ed8",style="<|-|>",ls="--"); ax.text(11.15,3.0,"Wi-Fi",fontsize=7,color="#1d4ed8",rotation=90)
ax.plot([],[],color=POW,lw=2,label="power"); ax.plot([],[],color=SIG,lw=1.2,label="signal / data")
ax.legend(loc="lower center",bbox_to_anchor=(0.5,-0.06),ncol=2,frameon=False,fontsize=7.5)
save(fig,"fig_block_diagram.png")

# Fig: project workflow / status
stages=[("Concept &\nresearch","d"),("Component\nselection","d"),("Component\nprocurement","d"),("PCB design\n& layout","d"),("Gerber\ngeneration","d"),
        ("PCB\nfabrication","n"),("3D printing","n"),("Mechanical\nassembly","t"),("Electronics\nintegration","t"),("Firmware &\ngait code","t"),("Testing &\ndemonstration","t")]
fig,ax=canvas(7.4,3.2,12.2,5.2)
col={"d":DONE,"n":NOW,"t":TODO}; fcs={"d":"#dcfce7","n":"#fef3c7","t":"#f3f4f6"}
pos=[]
for i in range(6): pos.append((0.1+i*2.02,3.2))
for i in range(5): pos.append((10.2-i*2.02,0.6))
for i,((t,s),(x,y)) in enumerate(zip(stages,pos)):
    box(ax,x,y,1.75,1.25,f"{i+1}. {t}",fc=fcs[s],ec=col[s],lw=1.5,fs=7)
for i in range(len(pos)-1):
    (x1,y1),(x2,y2)=pos[i],pos[i+1]
    if y1==y2:
        if x2>x1: arrow(ax,x1+1.75,y1+0.62,x2,y2+0.62)
        else: arrow(ax,x1,y1+0.62,x2+1.75,y2+0.62)
    else: arrow(ax,x1+0.87,y1,x2+0.87,y2+1.25)
ax.text(0.1,4.85,"■ completed",color=DONE,fontsize=8); ax.text(2.2,4.85,"■ immediate next steps",color=NOW,fontsize=8); ax.text(6.3,4.85,"■ later stages",color=TODO,fontsize=8)
ax.annotate("next stages",xy=(10.2+0.9,3.2+1.3),xytext=(8.9,4.85),fontsize=8,color=NOW,fontweight="bold",arrowprops=dict(arrowstyle="-|>",color=NOW))
save(fig,"fig_progress_path.png")

# Fig: power distribution detail
fig,ax=canvas(7.4,3.9,13.4,6.4)
box(ax,0.1,2.7,1.9,1.0,"Battery +\n(after fuse\n& switch)",fc="#fef2f2",ec=POW,fs=7.8)
rows=[("XL4016 → 6.00 V","J26 → 6 V zone A","9 right-leg sockets;\nvia D1 (1N4007) ≈ 5.2 V to both camera SG90s","C1 2200 µF, C6 470 µF"),
      ("XL4016 → 6.00 V","J27 → 6 V zone B","9 left-leg sockets","C2 2200 µF"),
      ("LM2596 → 5.00 V","J28 → +5 V zone","Uno 5V pin, PCA9685 logic,\n3 × HC-SR04, IMU, LEDs, buzzer","C3 470 µF, C7–C10 100 nF"),
      ("LM2596 → 5.00 V","J29 → +5 V CAM","ESP32-CAM only (J25 pin 1)","C4 1000 µF, C5 470 µF, C11 100 nF")]
for i,(a,b,c,d) in enumerate(rows):
    y=5.2-i*1.45
    box(ax,2.6,y,2.1,0.95,a,ec=POW,fs=7.6); box(ax,5.1,y,2.3,0.95,b,ec=POW,fs=7.2)
    box(ax,7.8,y,5.5,1.05,c+"\n"+d,fs=6.6,fc="#f8fafc")
    ax.plot([2.0,2.3,2.3],[3.2,3.2,y+0.47],color=POW,lw=1.5); arrow(ax,2.3,y+0.47,2.6,y+0.47,c=POW,lw=1.5)
    arrow(ax,4.7,y+0.47,5.1,y+0.47,c=POW,lw=1.5); arrow(ax,7.4,y+0.47,7.8,y+0.47,c=POW,lw=1.5)
ax.text(0.1,0.1,"Battery sense: J30 → R1 100 kΩ / R2 47 kΩ divider → A0 (C12 100 nF).  All negatives meet on the bottom ground pour (star point).",fontsize=7,color=SIG)
save(fig,"fig_power_rails.png")

# Fig: support polygon & SSM
fig,ax=plt.subplots(figsize=(6.4,3.2))
ax.set_aspect("equal"); ax.axis("off")
body=Polygon([(-0.6,-1.2),(0.6,-1.2),(0.8,0),(0.6,1.2),(-0.6,1.2),(-0.8,0)],closed=True,fc="#f1f5f9",ec=INK,lw=1)
ax.add_patch(body)
feet={"R1":(1.7,1.5),"R2":(1.9,0),"R3":(1.7,-1.5),"L1":(-1.7,1.5),"L2":(-1.9,0),"L3":(-1.7,-1.5)}
stance=["R1","L2","R3"]
tri=Polygon([feet[k] for k in stance],closed=True,fc="#dcfce7",ec=DONE,lw=1.5,alpha=0.8); ax.add_patch(tri)
hip={"R1":(0.6,1.2),"R2":(0.8,0),"R3":(0.6,-1.2),"L1":(-0.6,1.2),"L2":(-0.8,0),"L3":(-0.6,-1.2)}
for k,(x,y) in feet.items():
    ax.plot([hip[k][0],x],[hip[k][1],y],color=INK if k in stance else TODO,lw=1.3,ls="-" if k in stance else "--")
    ax.add_patch(Circle((x,y),0.11,fc=DONE if k in stance else "white",ec=INK))
    ax.text(x+(0.25 if x>0 else -0.25),y+0.18,k,fontsize=8,ha="center")
com=(0.15,0.05); ax.plot(*com,marker="x",color=POW,ms=9,mew=2); ax.text(com[0]+0.1,com[1]-0.3,"CoM\nprojection",fontsize=7.5,color=POW)
# distance to nearest edge (L2-R1 or L2-R3 or R1-R3)
import itertools
def dist(p,a,b):
    p,a,b=map(np.array,(p,a,b)); t=np.clip(np.dot(p-a,b-a)/np.dot(b-a,b-a),0,1); q=a+t*(b-a); return np.linalg.norm(p-q),q
ds=[dist(com,feet[a],feet[b]) for a,b in [("R1","L2"),("L2","R3"),("R3","R1")]]
d,q=min(ds,key=lambda z:z[0]); ax.annotate("",xy=q,xytext=com,arrowprops=dict(arrowstyle="<|-|>",color=POW,lw=1.2))
ax.text((com[0]+q[0])/2-0.55,(com[1]+q[1])/2+0.15,"$d_M$",color=POW,fontsize=10)
ax.text(2.6,0.8,"Stance legs (tripod): R1, L2, R3\nSwing legs: L1, R2, L3 (dashed)\n\nSupport polygon = triangle formed\nby the three stance feet.\nStatically stable if the CoM\nprojection lies inside it;\n$d_M$ = shortest distance to an edge.",fontsize=7.8,va="center")
ax.set_xlim(-2.3,6.2); ax.set_ylim(-1.9,1.9)
save(fig,"fig_support_polygon.png")

# Fig: leg geometry torque
fig,ax=canvas(6.4,2.8,10,4.4)
ax.add_patch(FancyBboxPatch((0.1,2.6),1.3,0.9,boxstyle="round,pad=0.02",fc="#f1f5f9",ec=INK)); ax.text(0.75,3.05,"BODY",ha="center",va="center",fontsize=8)
P=[(1.4,3.05),(2.3,3.05),(4.2,3.9),(4.9,0.6)]
for (a,b),lab in zip(zip(P[:-1],P[1:]),["coxa 30 mm","femur 65 mm","tibia 90 mm"]):
    ax.plot([a[0],b[0]],[a[1],b[1]],color=INK,lw=4,solid_capstyle="round")
    ax.text((a[0]+b[0])/2+(-1.35 if lab.startswith("tibia") else -0.9 if lab.startswith("femur") else -0.2),(a[1]+b[1])/2+(0.35 if lab.startswith("coxa") else 0.25),lab,fontsize=7.5)
for p,l in zip(P[:3],["hip servo","thigh servo","knee servo"]):
    ax.add_patch(Circle(p,0.13,fc="white",ec=POW,lw=1.5))
ax.text(1.2,2.55,"hip",fontsize=7,color=POW,ha="center"); ax.text(2.45,2.55,"thigh servo",fontsize=7,color=POW,ha="center"); ax.text(4.2,4.15,"knee servo",fontsize=7,color=POW,ha="center")
ax.plot([0,10],[0.47,0.47],color=TODO,lw=1)
ax.annotate("",xy=(4.9,0.25),xytext=(2.3,0.25),arrowprops=dict(arrowstyle="<|-|>",color=POW)); ax.text(3.6,0.0,"L ≤ 70 mm (horizontal)",fontsize=7.5,color=POW,ha="center")
ax.text(6.1,2.2,"Planned design check (build plan, Sec. 7):\n\nEstimated robot mass ≈ 2.22 kg\nTripod gait → 3 legs share the load\nLoad per leg ≈ 0.74 kg\nThigh torque ≈ 0.74 kg × 7.0 cm ≈ 5.2 kgf·cm\nMG996R stall ≈ 11 kgf·cm at 6 V\n→ about 47 % of stall (target < 60 %)\n\nThese are design estimates, not measurements.",fontsize=7.5,va="center",bbox=dict(fc="#f8fafc",ec=TODO))
save(fig,"fig_leg_torque.png")

# Fig: command flow
fig,ax=canvas(7.2,1.7,12.4,2.6)
items=["Phone\nweb page","ESP32-CAM\n(peripheral)","Arduino Uno R3\ngait algorithm","2 × PCA9685\nI²C, 50 Hz PWM","Servo sockets\non PCB","MG996R /\nSG90 joints"]
for i,t in enumerate(items):
    box(ax,0.1+i*2.07,1.0,1.7,1.1,t,fs=6.6,fc="#eff6ff" if i<2 else "white")
    if i: arrow(ax,0.1+i*2.07-0.37,1.55,0.1+i*2.07,1.55)
labs=["Wi-Fi","1 char, 9600 baud","I²C (A4/A5)","PWM signal","signal + 6 V"]
for i,l in enumerate(labs): ax.text(1.8+i*2.07+0.18,2.25,l,fontsize=6.5,ha="center",color=SIG)
ax.text(6.2,0.35,"Sensor feedback: 3 × HC-SR04 (D2–D7) and MPU6050 (I²C 0x68) report to the Uno, which can override a user command.",fontsize=7,ha="center",color=SIG)
save(fig,"fig_command_flow.png")

# Fig: 3D printing workflow
fig,ax=canvas(7.2,2.0,12.4,3.0)
st=["Study / adapt\nleg & body\nmodels","Tolerance cube\n+ one servo\nbracket","Check real\nMG996R fit\n(calipers)","Print & dry-fit\none full leg","Print remaining\n5 legs","Box, lid, plate,\ncamera mount,\ntrays"]
for i,t in enumerate(st):
    box(ax,0.1+i*2.07,1.1,1.8,1.3,t,fs=6.4,fc="#fef3c7" if i<2 else "#f3f4f6",ec=NOW if i<2 else TODO)
    if i: arrow(ax,0.1+i*2.07-0.27,1.75,0.1+i*2.07,1.75)
ax.annotate("",xy=(0.1+2*2.07+0.9,1.1),xytext=(0.1+3*2.07+0.9,1.1),arrowprops=dict(arrowstyle="-|>",color=POW,connectionstyle="arc3,rad=-0.5",ls="--"))
ax.text(5.6,0.25,"fit fails → adjust model and re-print before moving on",fontsize=7,color=POW,ha="center")
save(fig,"fig_print_workflow.png")

# Fig: PCB functional zones (conceptual)
fig,ax=canvas(6.4,4.2,11.4,8.4)
ax.add_patch(FancyBboxPatch((0.2,0.2),11,8.0,boxstyle="round,pad=0,rounding_size=0.2",fc="#ecfdf5",ec=INK,lw=1.5))
ax.text(5.7,7.8,"Conceptual functional grouping (not the actual copper layout) — approx. 114 × 84 mm",ha="center",fontsize=7.5,style="italic")
box(ax,0.5,5.3,3.3,2.1,"6 V zone A\nJ26 input, C1\nright-leg sockets\nJ31–J34, J5–J9\nD1 → J19, J20 (camera)",ec=POW,fs=7)
box(ax,0.5,2.9,3.3,2.1,"6 V zone B\nJ27 input, C2\nleft-leg sockets\nJ10–J18",ec=POW,fs=7)
box(ax,4.2,4.4,3.0,3.0,"PCA9685 U2 (0x40)\n\nPCA9685 U3 (0x41)\n\n(signal + GND rows\nonly; V+ row on\nno net)",fs=7)
box(ax,4.2,0.6,3.0,3.4,"Arduino Uno\nstacking headers\nJ1–J4\n(Uno sits underneath)",fs=7,fc="#f1f5f9")
box(ax,7.6,4.6,3.3,2.8,"+5 V logic zone\nJ28, C3, C7–C10\nJ21–J23 HC-SR04\nJ24 MPU6050 lead\nLEDs, BZ1 via Q1",ec=POW,fs=7)
box(ax,7.6,2.4,3.3,1.9,"+5 V CAM node\nJ29, C4, C5, C11\nJ25 camera + JP1\nR7/R8 divider",ec="#1d4ed8",fs=7)
box(ax,0.5,0.6,3.3,1.9,"Battery sense\nJ30, R1/R2, C12 → A0\nR3 pull-up on OE (D10)",fs=7)
box(ax,7.6,0.6,3.3,1.5,"Bottom layer:\nground pour = star point",fs=7,fc="#f8fafc")
save(fig,"fig_pcb_zones.png")
print("done")
