import { initializeApp } from 'firebase/app';
import { getFirestore, collection, addDoc, getDocs, query, where, deleteDoc } from 'firebase/firestore';
import firebaseConfig from './firebase-applet-config.json';

const app = initializeApp(firebaseConfig);
const db = getFirestore(app, firebaseConfig.firestoreDatabaseId);

const months = [
  "114年4月", "114年5月", "114年6月", "114年7月", "114年8月", "114年9月",
  "114年10月", "114年11月", "114年12月", "115年1月", "115年2月"
];

const data: Record<string, number[]> = {
  "01.急診大樓─總盤": [31489980, 32027779, 32683731, 33157820, 33842816, 34530080, 35243480, 35842560, 35970380, 36415379, 36843817],
  "02.急診大樓─後半段": [726729, 1078902, 1443578, 1835297, 2227741, 2634298, 3072253, 3434805, 3768358, 4072291, 4359073],
  "03.職務宿舍": [4823818, 5025438, 5199501, 5206243, 5212797, 5219268, 5226049, 5232292, 5238568, 5244720, 5250338],
  "04.門診大樓": [7181976, 7270009, 7378873, 7491604, 7588683, 7695419, 7810413, 7909772, 7996827, 8078616, 8155515],
  "05.行政大樓": [546418, 552379, 557941, 563224, 568629, 574130, 579981, 584784, 589200, 593431, 597277],
  "06.復健大樓": [500567, 509868, 518777, 527774, 537171, 546312, 555913, 562998, 569647, 575133, 580039],
  "07.役男宿舍": [520, 1175, 1829, 2421, 3018, 3685, 3919, 4120, 4321, 4487, 4583],
  "09.動力中心": [901320, 921371, 938641, 951102, 966666, 982230, 997832, 1010215, 1024269, 1038450, 1053026],
  "10.水塔": [659127, 675957, 691221, 706145, 720589, 736780, 754097, 770539, 789007, 806626, 822369],
  "11.汙水處理廠": [6668, 6708, 6751, 6789, 6875, 6970, 7063, 7148, 7326, 7259, 7344],
  "12.廚房": [116830, 119128, 121140, 121185, 121232, 121285, 121330, 121971, 121417, 121462, 121502],
  "13.精神科大樓": [215233, 218923, 222481, 225905, 229419, 232950, 236464, 239419, 242408, 245256, 247871],
  "14.AB棟": [12311076, 12568870, 12805104, 13057062, 13341401, 13619265, 13915587, 14158976, 14390916, 14606268, 14813689],
  "15.松柏園": [7189639, 7298079, 7404398, 7516000, 7624878, 7729148, 7838023, 7930198, 8020161, 8105055, 8187358],
  "16.廢棄物處理廠": [604, 1198, 1759, 2333, 2886, 3444, 4013, 4491, 4936, 5338, 5735],
  "17.懷遠堂": [5681.9, 13013, 18513, 19770, 21236, 23787, 26317, 28260, 32742, 33660, 34605]
};

const adjustments: Record<string, Record<string, number>> = {
  "01.急診大樓─總盤": { "114年12月": 431584 },
  "11.汙水處理廠": { "115年1月": 67 }
};

async function seed() {
  console.log("Starting seeding...");
  const readingsCol = collection(db, 'readings');

  // Optional: Clear existing data
  const snapshot = await getDocs(readingsCol);
  console.log(`Deleting ${snapshot.size} existing documents...`);
  for (const doc of snapshot.docs) {
    await deleteDoc(doc.ref);
  }

  for (const meter of Object.keys(data)) {
    const values = data[meter];
    for (let i = 0; i < values.length; i++) {
      const date = months[i];
      const value = values[i];
      const prevValue = i > 0 ? values[i - 1] : 0;
      const usage = i > 0 ? value - prevValue : 0;
      const adjustment = (adjustments[meter] && adjustments[meter][date]) || 0;

      // Create a base timestamp for sorting (Year * 12 + Month)
      const mMatch = date.match(/(\d+)年(\d+)月/);
      const ts = mMatch ? parseInt(mMatch[1]) * 12 + parseInt(mMatch[2]) : Date.now();

      await addDoc(readingsCol, {
        date,
        meter,
        value,
        adjustment,
        usage,
        ts
      });
      console.log(`Added ${meter} for ${date}`);
    }
  }
  console.log("Seeding completed!");
  process.exit(0);
}

seed().catch(err => {
  console.error(err);
  process.exit(1);
});
