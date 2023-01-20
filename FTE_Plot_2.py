import matplotlib.pyplot as plt
import numpy as np

x1 = np.array([8,8,20,28,36])
x2 = np.array([8,8,20,28,36])
#x3 = np.array([98,99,99,97,97])

y1 = np.array([300,600,1200,1800,2400])

plt.plot(x1, y1, label = "Current FTE", linestyle="-.")
plt.plot(x2, y1, label = "FTE Saved", linestyle=":")
#plt.plot(x3, y1, label = "Bot Efficiency", linestyle=":")
plt.ylabel('Number of Placements',color = 'maroon',fontsize = '14.0')
plt.xlabel('Current FTE = [Trafficker + QA]',color = 'darkblue',fontsize = '14.0')
plt.title('Sizmek Workflow Current FTE vs FTE Saved',color='darkgreen',fontsize = '20.0')

#Include Text
props = dict(boxstyle='round', facecolor='wheat', alpha=0.5)
plt.text(14.5,634.0, "*FTE saved by bot is approximately the "
                      "\nsame FTE consumed manually"
                      "\n*Workflow execution by bot and user will be under a minute"
         ,fontsize =8.0,color = 'forestgreen',bbox= props, wrap = True)

plt.legend()
plt.show()