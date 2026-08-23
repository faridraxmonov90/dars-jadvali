import { initializeApp } from "firebase/app";
import { getFirestore, doc, getDoc, setDoc } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyBpxrOBn4mYAG2SnB9E2eLUStXyVahc-k0",
  authDomain: "dars-jadvali-532c9.firebaseapp.com",
  projectId: "dars-jadvali-532c9",
  storageBucket: "dars-jadvali-532c9.firebasestorage.app",
  messagingSenderId: "927524281167",
  appId: "1:927524281167:web:1588c387951a4aff7c833a",
};

const app = initializeApp(firebaseConfig);
const dbFirestore = getFirestore(app);

export async function storageGet(key) {
  const snap = await getDoc(doc(dbFirestore, "storage", key));
  return snap.exists() ? { value: snap.data().value } : null;
}

export async function storageSet(key, value) {
  await setDoc(doc(dbFirestore, "storage", key), { value });
  return { value };
}