import matplotlib.pyplot as plt
import numpy as np

fig = plt.figure(figsize=(11,8))
ax1 = fig.add_subplot(111)
#BRD Time consumed [minutes]
Workflow_Time= (60,480,180,180,180)
#Third BRD is 50 creatives and Fourth BRD is 25 partners in Campaigns
Bot_Time  = (1,1,1,5,5)




# multiple line plot
ax1.plot(Workflow_Time,Bot_Time,marker='o', markerfacecolor='blue', markersize=12, color='skyblue', linewidth=4,label="")


#Include Text
props = dict(boxstyle='round', facecolor='wheat', alpha=0.5)
plt.text(4.3,50.0, "*Based on BRD Items Time consumed"
                   "\n*BRD-1,2,3 are based on Placements"
                    "\n*Third BRD is 50 creatives"
                   "\n*Fourth BRD is 25 partners in Campaigns"
         ,fontsize = 12.5,color = 'forestgreen',bbox= props, wrap = True)
#Title and labels
plt.ylabel('Time consumed by Bot[min]',color = 'maroon',fontsize = '14.0')
plt.xlabel('Manual workflow time consumed[min]',color = 'darkblue',fontsize = '14.0')
plt.title('Sizmek Workflow Manual vs Bot',color='darkgreen',fontsize = '20.0')
plt.legend()

#ax1.axhline(y=28.0,xmin=0.0, xmax=1.0, color='r')
#ax1.axhline(y=15.0,xmin=0.0, xmax=1.0, color='r')


plt.show()

