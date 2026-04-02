const { onCustomEventPublished } = require("firebase-functions/v2/eventarc");
const { initializeApp } = require("firebase-admin/app");
const { getFirestore } = require("firebase-admin/firestore");

initializeApp();

// Firebase Authentication のユーザー削除時に Firestore の users ドキュメントも削除
exports.onUserDeleted = onCustomEventPublished(
  "google.firebase.authentication.user.v1.deleted",
  async (event) => {
    const uid = event.data.uid || event.subject.replace("users/", "");
    try {
      await getFirestore().collection("users").doc(uid).delete();
      console.log(`Firestore user document deleted: ${uid}`);
    } catch (error) {
      console.error(`Error deleting Firestore document for ${uid}:`, error);
    }
  }
);
