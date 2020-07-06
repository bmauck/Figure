from tools.tools import *
from tools.email import *
import glob

df_fb = pd.DataFrame()
for f in glob.glob(data_dir + 'orig\\assets\\heloc\\originated_assets-*.csv'):
#     print(f)
    df = pd.read_csv(f)
    df['cutoff_date'] = df['cutoff_date'][1]
    df = clean_and_bin_tape(df)    
    df = df[['loan_id', 'cutoff_date', 'prin_bal', 'loan_status', 'orig_date']]
    df_fb = pd.concat([df_fb, df])

df_fb = df_fb.merge(get_owners())

df_fb_owner = pd.DataFrame(
    df_fb[df_fb['loan_status'] == 'FORBEARANCE'].groupby(['owner', 'cutoff_date'])['prin_bal'].sum() 
    / df_fb.groupby(['owner', 'cutoff_date'])['prin_bal'].sum()
    )
df_fb_port = pd.DataFrame(
    df_fb[df_fb['loan_status'] == 'FORBEARANCE'].groupby(['cutoff_date'])['prin_bal'].sum() 
    / df_fb.groupby(['cutoff_date'])['prin_bal'].sum()
    )

df_fb_owner = df_fb_owner.reset_index()
df_fb_owner = df_fb_owner.rename(columns={'prin_bal':'%_fb'})
df_fb_port = df_fb_port.rename(columns={'prin_bal':'%_fb'})

df_fb_owner['\u0394'] = df_fb_owner['%_fb'].diff()
df_fb_port['\u0394'] = df_fb_port['%_fb'].diff()

df_fb.to_csv(data_dir + 'forbearance\\rpt_forbearance-{}.csv'.format(tday))

fig, (ax1, ax2, ax3) = plt.subplots(1, 3, sharey=True, figsize=(25,10))
fig.suptitle('Forbearance Statistics {}'.format(tday), fontsize=18)
ax1.plot(df_fb_port, marker='o')
ax1.set_title('Portfolio', fontsize=14)
ax1.legend(df_fb_port.columns)
ax1.grid(linestyle='-', linewidth='.5', color='black')
ax1.table(cellText=df_fb_port.values.round(3)
          ,rowLabels=df_fb_port.index
          ,colLabels=['% fb', '\u0394']
          ,colWidths=[0.3,0.3]
          ,colLoc='center'
          ,loc='bottom'
          ,bbox=[0.5, -0.7, .3, .6])

df_plot = df_fb_owner[df_fb_owner['owner'] == 'FLOC 2020-1']
df_plot = df_plot.set_index('cutoff_date')
df_plot.drop('owner', inplace=True, axis=1)
df_plot.iloc[0,1] = 0
ax2.plot(df_plot, marker='o')
ax2.set_title('FLOC 2020-1', fontsize=14)
ax2.legend(df_fb_port.columns)
ax2.grid(linestyle='-', linewidth='0.5', color='black')
ax2.table(cellText=df_plot.values.round(3)
          ,rowLabels=df_plot.index
          ,colLabels=['% fb', '\u0394']
          ,colWidths=[0.3, 0.3]
          ,colLoc='center'
          ,loc='bottom'
          ,fontSize=6
          ,bbox=[0.5, -0.7, 0.3, .6])

df_plot = df_fb_owner[df_fb_owner['owner'] == 'VFN']
df_plot = df_plot.set_index('cutoff_date')
df_plot.drop('owner', inplace=True, axis=1)
df_plot.iloc[0,1] = 0
ax3.plot(df_plot, marker='o')
ax3.set_title('VFN', fontsize=14)
ax3.legend(df_fb_port.columns)
ax3.grid(linestyle='-', linewidth='0.5', color='black')
ax3.table(cellText=df_plot.values.round(3)
          ,rowLabels=df_plot.index
          ,colLabels=['% fb', '\u0394']
          ,colWidths=[0.3, 0.3]
          ,colLoc='center'
          ,loc='bottom'
          ,fontSize=6
          ,bbox=[0.5, -0.7, 0.3, .6])

plt.tight_layout(pad=3, rect=[1, 0.3, 0.95, .95])

fig.canvas.draw()
plt.savefig(data_dir + 'forbearance\\rpt_forbearance-{}.png'.format(tday), bbox_inches='tight')

send_email(
    recipient='capitalmarkets@figure.com'
    ,subject='Forbearance Report {}'.format(tday)
    ,contents=[yagmail.inline(data_dir + 'forbearance\\rpt_forbearance-{}.png'.format(tday))]
    ,cc='bmauck@figure.com'
    )