import { initializeApp } from "firebase/app";
import { getFirestore, doc, getDoc, setDoc } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyD4acAPk8nH5BzGi7ESXX-P46UJeBO1al0",
  authDomain: "megastudy01.firebaseapp.com",
  projectId: "megastudy01",
  storageBucket: "megastudy01.firebasestorage.app",
  messagingSenderId: "982234390958",
  appId: "1:982234390958:web:cd39eeeef381054927283f"
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

export async function loadS(key) {
  try {
    const snap = await getDoc(doc(db, "storage", key));
    return snap.exists() ? snap.data().value : null;
  } catch { return null; }
}

export async function saveS(key, value) {
  try {
    await setDoc(doc(db, "storage", key), { value });
  } catch(e) { console.error(e); }
}