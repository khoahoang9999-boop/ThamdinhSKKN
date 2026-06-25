import { initializeApp } from 'firebase/app';
import { getFirestore } from 'firebase/firestore';
import { getAuth } from 'firebase/auth';

const firebaseConfig = {
  apiKey: "AIzaSyASigHThGWSkJnSUfyemnvnWrmdhh45_NE",
  authDomain: "thamdinhskkn.firebaseapp.com",
  projectId: "thamdinhskkn",
  storageBucket: "thamdinhskkn.firebasestorage.app",
  messagingSenderId: "270326182440",
  appId: "1:270326182440:web:31875d8b8c4441d206bb1d",
  measurementId: "G-JRYVNNKYL7"
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
export const auth = getAuth(app);
