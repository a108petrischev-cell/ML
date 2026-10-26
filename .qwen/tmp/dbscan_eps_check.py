import pandas as pd
import numpy as np
from sklearn.cluster import DBSCAN
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import silhouette_score

pd.set_option('display.width', 250)

data = pd.read_csv('/Users/alexeypetrischev/all_repos/OTUS/ML/DBscan/german_credit_data.csv', sep=',')
df = data.drop(columns=['Unnamed: 0'])
df['Saving accounts'] = df['Saving accounts'].fillna('unknown')
df = df.drop(columns=['Checking account'])
df = pd.get_dummies(df, columns=['Sex', 'Job', 'Housing', 'Saving accounts', 'Purpose'], drop_first=True)

X = StandardScaler().fit_transform(df)

# sweep on X (standardized), min_samples=20, like the fixed notebook will
print("DBSCAN on X, min_samples=20:")
for eps in np.round(np.arange(0.5, 5.0, 0.25), 2):
    m = DBSCAN(eps=eps, min_samples=20).fit(X)
    lab = m.labels_
    n_cl = len(np.unique(lab[lab != -1]))
    n_noise = int((lab == -1).sum())
    sil = silhouette_score(X, lab) if n_cl > 1 else 0.0
    sizes = sorted(pd.Series(lab[lab != -1]).value_counts().tolist(), reverse=True) if n_cl else []
    print(f"  eps={eps:.2f}: clusters={n_cl}, noise={n_noise}, sil={sil:.3f}, top_sizes={sizes[:8]}")

raw = data.drop(columns=['Unnamed: 0']).copy()
raw['Saving accounts'] = raw['Saving accounts'].fillna('unknown')
raw = raw.drop(columns=['Checking account'])

def profile(r):
    num = r[['Age', 'Credit amount', 'Duration']].mean().round(1)
    cats = r[['Sex', 'Job', 'Housing', 'Saving accounts', 'Purpose']].agg(lambda s: s.mode().iloc[0])
    return (f"n={len(r)}, Age={num['Age']}, Credit={num['Credit amount']}, Dur={num['Duration']}, "
            f"Sex={cats['Sex']}, Job={cats['Job']}, Housing={cats['Housing']}, Sav={cats['Saving accounts']}, Purp={cats['Purpose']}")

# elbow eps = median of 20-NN on X
from sklearn.neighbors import NearestNeighbors
nn = NearestNeighbors(n_neighbors=20); nn.fit(X)
d, _ = nn.kneighbors(X)
eps = float(np.median(np.sort(d[:, -1], axis=0)))
print(f"\nElbow eps (median 20-NN on X) = {eps:.3f}")

lab = DBSCAN(eps=eps, min_samples=20).fit(X).labels_
r2 = raw.copy(); r2['cluster'] = lab
print("\n=== Final: DBSCAN(eps=elbow, min_samples=20) on X ===")
print(r2['cluster'].value_counts().sort_index().to_string())
for c, r in r2[r2['cluster'] != -1].groupby('cluster'):
    print(f"  cluster {c}: " + profile(r))
nz = r2[r2['cluster'] == -1]
print(f"  NOISE: " + profile(nz))
sil = silhouette_score(X, lab) if len(np.unique(lab[lab != -1])) > 1 else 0.0
print(f"  silhouette: {sil:.3f}")
