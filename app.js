/* =====================================================
   JOIN Απογραφή — Main App
   ===================================================== */

// ==============================
// ITEMS DATA (192 items from Excel)
// ==============================
const ITEMS = [
  {row:2,code:null,name:"PROTEIN ALMOND 3KG (AMINO ANIMO)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:3,code:null,name:"PROTEIN PEA 3KG (AMINO ANIMO)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:4,code:null,name:"PROTEIN VANILLA 3KG (AMINO ANIMO)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:5,code:null,name:"PROTEIN CACAO 3KG (AMINO ANIMO)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:6,code:"50017Σ",name:"BAR BANANA ALMOND(BEIKERS)",unit_price:0,unit:"ΤΜΧ",category:"FOOD"},
  {row:7,code:"50016Σ",name:"BAR CHOCO STRAWBERRY(BEIKERS)",unit_price:0,unit:"ΤΜΧ",category:"FOOD"},
  {row:8,code:"50018Σ",name:"BAR HONEY CRANBERRY(BEIKERS)",unit_price:0,unit:"ΤΜΧ",category:"FOOD"},
  {row:9,code:"50014Σ",name:"BAR WHEY PROTEIN(BEIKES)",unit_price:0,unit:"ΤΜΧ",category:"FOOD"},
  {row:10,code:"ΜΠΡ-ΦΛΕΙ",name:"BROWN FLAKES (ΜΕΝΕΞΟΠΟΥΛΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:11,code:"ΣΠΟ-ΡΟΙ",name:"CHIA ΣΠΟΡΟΙ (MΕΝΕΞΟΠΟΥΛΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:12,code:"50032Σ",name:"COOKIE BROWNIE SOFT(BEIKERS)",unit_price:0,unit:"ΤΜΧ",category:"FOOD"},
  {row:13,code:"50034Σ",name:"COOKIE OAT(BEIKERS)",unit_price:0,unit:"ΤΜΧ",category:"FOOD"},
  {row:14,code:"50031Σ",name:"COOKIE PEANUT PROTEIN(BEIKERS)",unit_price:0,unit:"ΤΜΧ",category:"FOOD"},
  {row:15,code:"ΚΡΑ–ΜΠΕΡ",name:"CRANBERRIES (ΜΕΝΕΞΟΠΟΥΛΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:16,code:"A 0922",name:"FLAKES ΚΑΡΥΔΑΣ (ΑΡΜΟΝΙΑ ΓΕΥΣΗΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:17,code:"ΜΑΚ–ΑΣΚΟ",name:"MACA ΣΚΟΝΗ (ΜΕΝΕΞΟΠΟΥΛΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:18,code:"ΤΣΑΙ-MATCH",name:"MATCHA (ΜΕΝΕΞΟΠΟΥΛΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:22,code:"A 000907",name:"ACAI ΚΤΨ 5 KG (ΑΡΜΟΝΙΑ ΓΕΥΣΗΣ)",unit_price:0,unit:"TMX",category:"FOOD"},
  {row:23,code:"Χ 01-198-202-02",name:"VEGAN PROTEIN VANILLA (X-TREME STORES)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:24,code:"X 01-136-002-04",name:"PROTEIN WARRIORLAB WHEY(X-TREME STORES)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:25,code:"Α-023",name:"ΑΒΓΑ 30ΑΔΑ (ΤΣΑΚΙΡΗΣ)",unit_price:0,unit:"TMX",category:"FOOD"},
  {row:26,code:"A0605",name:"ΑΓΡΙΟΡΑΠΑΝΟ-HORSERADISH (ΑΡΜΟΝΙΑ ΓΕΥΣΗΣ)(140ΓΡ)",unit_price:0,unit:"ΤΜΧ",category:"FOOD"},
  {row:27,code:"001-0000",name:"ΑΛΑΤΙ ΧΟΝΤΡΟ/ΨΙΛΟ ΧΥΜΑ (ΜΕΝΕΞΟΠΟΥΛΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:28,code:null,name:"ΑΜΥΓΔΑΛΟΨΙΧΑ ΨΗΜΕΝΗ ΠΕΡΛΑ (ΜΕΝΕΞΟΠΟΥΛΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:29,code:"ΑΜΥ–ΩΑΜΕ",name:"ΑΜΥΓΔΑΛΟ ΩΜΟ (ΜΕΝΕΞΟΠΟΥΛΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:30,code:null,name:"ΒΑΝΙΛΙΝΗ ΕΞΩΤΕΡΙΚΟΥ ΣΚΟΝΗ (ΜΕΝΕΞΟΠΟΥΛΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:31,code:null,name:"ΒΟΥΤΥΡΟ ΦΥΣΤΙΚΙΟΥ ΑΙΓΙΝΗΣ 212ΓΡ (ΜΕΝΕΞΟΠΟΥΛΟΣ)",unit_price:0,unit:"ΤΜΧ",category:"FOOD"},
  {row:32,code:"Φ-024-2",name:"ΒΡΩΜΗ 1ΚΙΛ (ΤΣΑΚΙΡΗΣ)",unit_price:0,unit:"ΤΜΧ",category:"FOOD"},
  {row:19,code:null,name:"ΒΡΩΜΗ ΜΠΡΟΣΤΑ",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:33,code:"1915",name:"ΓΑΛΟΠΟΥΛΑ (ΑΡΜΟΝΙΑ ΓΕΥΣΗΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:34,code:"2818",name:"ΓΙΑΟΥΡΤΙ (ΜΕΒΓΑΛ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:35,code:"2308",name:"ΓΙΑΟΥΡΤΙ LACTOSE FREE (ΜΕΒΓΑΛ)",unit_price:0,unit:"ΤΜΧ",category:"FOOD"},
  {row:36,code:"Φ-016-2",name:"ΓΚΡΑΝΟΛΑ CHOCO (ΤΣΑΚΙΡΗΣ)",unit_price:0,unit:"ΤΜΧ",category:"FOOD"},
  {row:37,code:"A 000381",name:"ΕΛΑΙΟΛΑΔΟ 5ΛΙΤΡΑ (ΑΡΜΟΝΙΑ ΓΕΥΣΗΣ)",unit_price:0,unit:"ΤΜΧ",category:"FOOD"},
  {row:38,code:"1255",name:"ΕΛΙΕΣ ΡΟΔΕΛΑ 5.62ΚΙΛΑ (ΑΡΜΟΝΙΑ ΓΕΥΣΗΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:39,code:"ΗΛΙ-ΨΙΕΞ",name:"ΗΛΙΟΣΠΟΡΟΣ (ΜΕΝΕΞΟΠΟΥΛΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:40,code:"A 000651",name:"ΙΟΝ ΣΤΑΓΟΝΕΣ 2ΚΙΛΑ (ΑΡΜΟΝΙΑ ΓΕΥΣΗΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:41,code:"ΚΑΚ–ΑΟΚΟ",name:"ΚΑΚΑΟ ΚΟΜΜΑΤΙΑ (ΜΕΝΕΞΟΠΟΥΛΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:42,code:"A 000846",name:"KAKAO ΣΚΟΝΗ ΙΟΝ (ΑΡΜΟΝΙΑ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:43,code:"A 001460",name:"ΚΑΛΑΜΠΟΚΙ ΣΠΟΡΟΙ ΚΟΝΣΕΡΒΑ (ΑΡΜΟΝΙΑ ΓΕΥΣΗΣ)",unit_price:0,unit:"ΤΜΧ",category:"FOOD"},
  {row:44,code:"ΚΑΝ-ΣΚΟΒ",name:"ΚΑΝΕΛΑ (MΕΝΕΞΟΠΟΥΛΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:45,code:"ΚΑΡ–1ΒΟΥ",name:"ΚΑΡΥΔΙΑ (ΜΕΝΕΞΟΠΟΥΛΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:46,code:"ΚΙΝ–ΟΑΒΙ",name:"ΚΙΝΟΑ (ΜΕΝΕΞΟΠΟΥΛΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:47,code:"A 020015",name:"ΚΟΤΟΠΟΥΛΟ (DELICES)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:48,code:"A 0849",name:"ΛΕΜΟΝΟΠΙΠΕΡΟ (ΑΡΜΟΝΙΑ ΓΕΥΣΗΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:49,code:null,name:"ΜΑΚΑΝΤΑΜΙΑ ΩΜΑ (ΜΕΝΕΞΟΠΟΥΛΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:50,code:"1605",name:"ΜΕΛΙ (FOODWISE)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:51,code:"A 001788",name:"ΜΟΤΣΑΡΕΛΑ 125ΓΡ (ΑΡΜΟΝΙΑ ΓΕΥΣΗΣ)",unit_price:0,unit:"ΤΜΧ",category:"FOOD"},
  {row:52,code:"2340",name:"ΜΟΥΣΤΑΡΔΑ 4ΚΙΛΑ (ΑΡΜΟΝΙΑ ΓΕΥΣΗΣ)",unit_price:0,unit:"ΤΜΧ",category:"FOOD"},
  {row:53,code:"A 021009",name:"ΞΥΔΙ ΑΠΛΟ 2ΛΙΤΡΑ (ΑΡΜΟΝΙΑ ΓΕΥΣΗΣ)",unit_price:0,unit:"ΤΜΧ",category:"FOOD"},
  {row:54,code:"A 0551",name:"ΠΑΝΤΖΑΡΙ ΒΡΑΣΜΕΝΟ (ΑΡΜΟΝΑ) 500ΓΡ",unit_price:0,unit:"ΤΜΧ",category:"FOOD"},
  {row:55,code:"A 020030",name:"ΠΑΡΜΕΖΑΝΑ ΦΛΥΔΕΣ 500ΓΡ (ΑΡΜΟΝΙΑ ΓΕΥΣΗΣ)",unit_price:0,unit:"ΤΜΧ",category:"FOOD"},
  {row:56,code:"ΠΙΠ-ΕΡΙ",name:"ΠΙΠΕΡΙ ΜΑΥΡΟ ΣΚΟΝΗ (ΜΕΝΕΞΟΠΟΥΛΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:57,code:null,name:"ΠΙΠΕΡΙ ΜΑΥΡΟ ΣΠΥΡΙ (MΕΝΕΞΟΠΟΥΛΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:58,code:"A 001709",name:"ΠΡΟΣΟΥΤΟ 500ΓΡ (ΑΡΜΟΝΙΑ ΓΕΥΣΗΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:59,code:"ΑΓΚ-ΑΒΕ",name:"ΣΙΡΟΠΙ ΑΓΑΥΗΣ (ΜΕΝΕΞΟΠΟΥΛΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:60,code:"A 000091",name:"ΣΟΛΟΜΟΣ (ΑΡΜΟΝΙΑ ΓΕΥΣΗΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:61,code:"ΣΤΑ-ΤΟΥΡ",name:"ΣΤΑΦΙΔΑ (ΜΕΝΕΞΟΠΟΥΛΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:62,code:"629",name:"ΤΑΧΙΝΙ ΟΛΙΚΗΣ (FOODWISE)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:63,code:"A 001028",name:"ΤΥΡΙ GOUDA LIGHT ΦΕΤΕΣ 200ΓΡ. (ΑΡΜΟΝΙΑ ΓΕΥΣΗΣ)",unit_price:0,unit:"ΤΜΧ",category:"FOOD"},
  {row:64,code:"A 000603",name:"ΤΥΡΙ ΚΡΕΜΑ 1,5ΚΙΛ (ΑΡΜΟΝΙΑ ΓΕΥΣΗΣ)",unit_price:0,unit:"ΤΜΧ",category:"FOOD"},
  {row:65,code:"811",name:"ΦΟΥΝΤΟΥΚΟΒΟΥΤΥΡΟ (FOODWISE)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:66,code:"ΦΥΣ–ΒΟΥΤ3",name:"ΦΥΣΤΙΚΟΒΟΥΤΥΡΟ (ΜΕΝΕΞΟΠΟΥΛΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:67,code:"Ψ 01-3143",name:"ΨΩΜΙ ΟΛΙΚΗΣ (DELICES)",unit_price:0,unit:"ΤΜΧ",category:"FOOD"},
  {row:68,code:"71.1002.9507",name:"ΦΕΤΑ ΧΩΡΙΑΤΙΚΟ ΤΟΣΤ (ΡΟΔΟΥΛΑ)",unit_price:0,unit:"ΤΜΧ",category:"FOOD"},
  {row:69,code:"BAKE-442",name:"ΨΩΜΙ ΧΩΡΙΣ ΓΛΟΥΤΕΝΗ (BRIOCHE)",unit_price:0,unit:"ΤΜΧ",category:"FOOD"},
  {row:70,code:null,name:"ΜΟΥΣ ΦΙΛΑΔΕΛΦΙΑ",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:71,code:null,name:"ΜΟΥΣ ΑΒΟΚΑΝΤΟ",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:72,code:null,name:"ΜΟΥΣ ΠΕΣΤΟ ΒΑΣΙΛΙΚΟΥ",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:73,code:null,name:"ΣΩΣ DRESSING",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:74,code:null,name:"PROTEIN PORRIDGE",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:75,code:null,name:"GREEN PORRIDGE",unit_price:0,unit:"ΚΙΛΟ",category:"FOOD"},
  {row:76,code:null,name:"PROTEIN BALL",unit_price:0,unit:"TMX",category:"FOOD"},
  // COFFEE
  {row:80,code:"A 0610",name:"ΕΒΑΠΟΡΕ 500ΓΡ. (ΑΡΜΟΝΙΑ ΓΕΥΣΗΣ)",unit_price:0,unit:"ΤΜΧ",category:"COFFEE"},
  {row:81,code:"1026",name:"ΚΑΨΟΥΛΑ ΕΛΛΗΝΙΚΟΥ COSTA COFFEE",unit_price:0,unit:"TMX",category:"COFFEE"},
  {row:82,code:"A 0658",name:"ΖΑΧΑΡΗ ΚΑΣΤΑΝΗ 1ΚΙΛΟ (ΑΡΜΟΝΙΑ ΓΕΥΣΗΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"COFFEE"},
  {row:83,code:"A 0447",name:"ΖΑΧΑΡΗ ΛΕΥΚΗ 1ΚΙΛΟ (ΑΡΜΟΝΙΑ ΓΕΥΣΗΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"COFFEE"},
  {row:84,code:"1002",name:"KAΦΕΣ COSTA COFFEE 1 ΚΙΛΟ",unit_price:0,unit:"ΚΙΛΟ",category:"COFFEE"},
  {row:85,code:"1005",name:"ΚΑΦΕΣ COSTA COFFEE DECAF. ΣΠΥΡΙ 500ΓΡ",unit_price:0,unit:"ΤΜΧ",category:"COFFEE"},
  {row:86,code:"13-50-004",name:"ΝΕΡΑ AΥΡΑ 500ML (ΑΜΠΑΤΖΗΣ)",unit_price:0,unit:"ΤΜΧ",category:"COFFEE"},
  {row:87,code:"10273",name:"ΝΕΣ DECAFF. (ΖΗΚΟΣ) 200ΓΡ",unit_price:0,unit:"ΤΜΧ",category:"COFFEE"},
  {row:88,code:null,name:"ΝΕΣ ΚΑΦΕ (ΜΑΝΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"COFFEE"},
  {row:89,code:"2012",name:"ΣΟΚΟΛΑΤΑ COSTA COFFEE",unit_price:0,unit:"ΚΙΛΟ",category:"COFFEE"},
  {row:90,code:"A 0667",name:"ΣΤΙΚ ΜΕΛΙ 150ΤΜΧ (ΑΡΜΟΝΙΑ)",unit_price:0,unit:"ΚΟΥΤΙ",category:"COFFEE"},
  {row:91,code:"A 0727",name:"ΣΤΙΚ ΖΑΧΑΡΙΝΗ (ΑΡΜΟΝΙΑ)",unit_price:0,unit:"ΚΟΥΤΙ",category:"COFFEE"},
  {row:92,code:"A 0671",name:"ΣΤΙΚ ΣΤΕΒΙΑ 500TMX (ΑΡΜΟΝΙΑ)",unit_price:0,unit:"ΚΟΥΤΙ",category:"COFFEE"},
  {row:93,code:null,name:"ΣΤΙΚ ΚΑΣΤΑΝΗ ΖΑΧΑΡΗ 1000ΤΜΧ COSTA COFFEE",unit_price:0,unit:"ΚΟΥΤΙ",category:"COFFEE"},
  {row:94,code:null,name:"ΣΤΙΚ ΛΕΥΚΗ ΖΑΧΑΡΗ 1000ΤΜΧ COSTA COFFEE",unit_price:0,unit:"ΚΟΥΤΙ",category:"COFFEE"},
  {row:95,code:null,name:"VERGNANO BLEND ΦΙΛΤΡΟΥ (COSTA COFFEE)",unit_price:0,unit:"ΤΜΧ",category:"COFFEE"},
  {row:96,code:null,name:"ΚΑΘΑΡΙΣΤΙΚΟ ΜΥΛΩΝ ΑΛΕΣΗΣ (COSTA COFFEE)",unit_price:0,unit:"ΚΟΥΤΙ",category:"COFFEE"},
  // JUICES
  {row:100,code:"M 000068",name:"BERRIES (ΑΡΜΟΝΙΑ ΓΕΥΣΗΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  {row:101,code:"M 0106",name:"ICEBERG (ΜΕΣΟΓΕΙΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  {row:102,code:"M 0104",name:"LIME (ΜΕΣΟΓΕΙΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  {row:103,code:"A 021007",name:"MANGO (ΑΡΜΟΝΙΑ ΓΕΥΣΗΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  {row:104,code:"M 0089",name:"ΑΒΟΚΑΝΤΟ (ΜΕΣΟΓΕΙΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  {row:105,code:"M 0014",name:"ΑΓΓΟΥΡΙ (ΜΕΣΟΓΕΙΟΣ)",unit_price:0,unit:"ΤΜΧ",category:"JUICES"},
  {row:106,code:"M 0083",name:"ΑΚΤΙΝΙΔΙΑ (ΜΕΣΟΓΕΙΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  {row:107,code:"M 0078",name:"ΑΝΑΝΑΣ (ΜΕΣΟΓΕΙΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  {row:108,code:"M 0217",name:"ΑΝΗΘΟ (ΜΕΣΟΓΕΙΟΣ)",unit_price:0,unit:"ΤΜΧ",category:"JUICES"},
  {row:109,code:"M 0023",name:"ΒΑΣΙΛΙΚΟ (ΜΕΣΟΓΕΙΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  {row:110,code:"M 0128",name:"ΒΕΛΑΝΙΔΙΑ (ΜΕΣΟΓΕΙΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  {row:111,code:"1299",name:"ΓΑΛΑ SPECIALISTAS 1.5% 1L (ΜΕΒΓΑΛ)",unit_price:0,unit:"ΤΜΧ",category:"JUICES"},
  {row:112,code:"1298",name:"ΓΑΛΑ SPECIALISTAS 3.5% 1LT (ΜΕΒΓΑΛ)",unit_price:0,unit:"ΤΜΧ",category:"JUICES"},
  {row:113,code:"0304-0100",name:"ΓΑΛΑ ΑΜΥΓΔΑΛΟΥ (KAFEA TERRA)",unit_price:0,unit:"ΤΜΧ",category:"JUICES"},
  {row:114,code:"0304-0101",name:"ΓΑΛΑ ΚΑΡΥΔΑΣ (KAFEA TERRA)",unit_price:0,unit:"ΤΜΧ",category:"JUICES"},
  {row:115,code:"0304-0102",name:"ΓΑΛΑ ΣΟΓΙΑΣ OAT (KAFEA TERRA)",unit_price:0,unit:"TMX",category:"JUICES"},
  {row:116,code:"M 0213",name:"ΔΥΟΣΜΟΣ (ΜΕΣΟΓΕΙΟΣ)",unit_price:0,unit:"ΤΜΧ",category:"JUICES"},
  {row:117,code:"M 0098",name:"ΚΑΡΟΤΑ (ΜΕΣΟΓΕΙΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  {row:118,code:"M 0371",name:"ΚΑΡΠΟΥΖΙ (ΜΕΣΟΓΕΙΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  {row:119,code:"M 0079",name:"ΜΗΛΑ ΖΑΓΟΡΙΝ (ΜΕΣΟΓΕΙΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  {row:120,code:"M 0080",name:"ΜΗΛΑ ΣΜΙΘ (ΜΕΣΟΓΕΙΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  {row:121,code:"M 0107",name:"ΜΠΑΝΑΝΑ (ΜΕΣΟΓΕΙΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  {row:122,code:"M 0091",name:"ΝΕΚΤΑΡΙΝΙ",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  {row:123,code:"M 0310",name:"ΝΤΟΜΑΤΑ (ΜΕΣΟΓΕΙΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  {row:124,code:"M 0785",name:"ΠΕΠΟΝΙ (ΜΕΣΟΓΕΙΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  {row:125,code:"M 0039",name:"ΠΙΠΕΡΙΑ ΦΛΩΡΙΝΗΣ (ΜΕΣΟΓΕΙΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  {row:126,code:"Μ 0125",name:"ΠΟΡΤΟΚΑΛΙΑ (ΜΕΣΟΓΕΙΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  {row:127,code:"0501-0001",name:"ΣΙΡΟΠΙ ROUTIN CARAMEL 1L (KAFEA TERRA)",unit_price:0,unit:"ΤΜΧ",category:"JUICES"},
  {row:128,code:"0501-0003",name:"ΣΙΡΟΠΙ ROUTIN HAZELNUT 1LT (KAFEA TERRA)",unit_price:0,unit:"ΤΜΧ",category:"JUICES"},
  {row:129,code:"M 0219",name:"ΣΠΑΝΑΚΙ ΒΑΒΥ (ΜΕΣΟΓΕΙΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  {row:130,code:"M 0045",name:"ΣΠΑΝΑΚΙ ΒΑΣΙΛΙΚΩΝ (ΜΕΣΟΓΕΙΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  {row:131,code:null,name:"ΣΤΑΦΥΛΙ (ΜΕΣΟΓΕΙΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  {row:132,code:"M 0096",name:"ΤΖΙΝΤΖΕΡ (ΜΕΣΟΓΕΙΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  {row:133,code:"A 020008",name:"ΦΡΑΟΥΛΑ ΚΤΨ (ΑΡΜΟΝΙΑ ΓΕΥΣΗΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  {row:134,code:"Μ 0129",name:"ΦΡΑΟΥΛΕΣ ΜΑΝΩΛΑΔΟΣ (ΜΕΣΟΓΕΙΟΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  {row:135,code:"3701",name:"ΧΥΜΟΣ ΛΕΜΟΝΙ 1Lt (ΑΡΜΟΝΙΑ)",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  {row:136,code:"500449",name:"ΑΒΟΚΑΝΤΟ ΦΡΟΥΤΟ ΚΤΨ (ΑΡΜΟΝΙΑ ΓΕΥΣΗΣ)",unit_price:0,unit:"ΚΙΛΟ",category:"JUICES"},
  // ΑΝΑΛΩΣΙΜΑ
  {row:140,code:null,name:"GEL ΧΕΡΙΩΝ ΑΝΤΙΣΗΠΤ. 5L (ΓΙΑΝΝΗΣ)",unit_price:0,unit:"ΤΜΧ",category:"ΑΝΑΛΩΣΙΜΑ"},
  {row:141,code:null,name:"XΛΩΡΙΝΗ ΠΑΧΥΡΕΥΣΤΗ 4L (GnP)",unit_price:0,unit:"ΤΜΧ",category:"ΑΝΑΛΩΣΙΜΑ"},
  {row:142,code:null,name:"YΓΡΟ ΚΡΕΜΟΣΑΠΟΥΝΟ ΑΦΡΟΣ 4L (GnP)",unit_price:0,unit:"ΤΜΧ",category:"ΑΝΑΛΩΣΙΜΑ"},
  {row:143,code:null,name:"ΓΑΝΤΙΑ ΝΙΤΡΙΛΙΟΥ/LATEX ΜΑΥΡΑ 100ΤΜΧ (GnP)",unit_price:0,unit:"ΚΟΥΤΙ",category:"ΑΝΑΛΩΣΙΜΑ"},
  {row:144,code:null,name:"ΘΕΡΜΙΚΑ ΡΟΛΑ 57*40*13-15ΤΡ (NETWAYS)",unit_price:0,unit:"ΤΜΧ",category:"ΑΝΑΛΩΣΙΜΑ"},
  {row:145,code:null,name:"ΘΕΡΜΙΚΑ ΡΟΛΑ 80*80*13-70ΤΡ (NETWAYS)",unit_price:0,unit:"ΤΜΧ",category:"ΑΝΑΛΩΣΙΜΑ"},
  {row:146,code:null,name:"ΚΑΘΑΡΙΣΤΙΚΟ DRY 5KG(ΤΣΟΓΙΑΣ)",unit_price:0,unit:"ΤΜΧ",category:"ΑΝΑΛΩΣΙΜΑ"},
  {row:147,code:null,name:"ΚΑΘΑΡΙΣΤΙΚΟ SALT 5KG(ΤΣΟΓΙΑΣ)",unit_price:0,unit:"ΤΜΧ",category:"ΑΝΑΛΩΣΙΜΑ"},
  {row:148,code:null,name:"ΚΑΘΑΡΙΣΤΙΚΟ ΔΑΠΕΔΟΥ 4L (ΓΙΑΝΝΗΣ)",unit_price:0,unit:"ΤΜΧ",category:"ΑΝΑΛΩΣΙΜΑ"},
  {row:149,code:null,name:"ΚΑΘΑΡΙΣΤΙΚΟ ΠΙΑΤΩΝ 4L (ΓΙΑΝΝΗΣ)",unit_price:0,unit:"ΤΜΧ",category:"ΑΝΑΛΩΣΙΜΑ"},
  {row:150,code:null,name:"ΚΑΘΑΡΙΣΤΙΚΟ ΠΛΥΝΤ. ΠΙΑΤ. 12KG (ΤΣΟΓΙΑΣ)",unit_price:0,unit:"ΤΜΧ",category:"ΑΝΑΛΩΣΙΜΑ"},
  {row:151,code:null,name:"ΚΑΘΑΡΙΣΤΙΚΟ ΤΖΑΜΙΩΝ 4L (ΓΙΑΝΝΗΣ)",unit_price:0,unit:"ΤΜΧ",category:"ΑΝΑΛΩΣΙΜΑ"},
  {row:152,code:null,name:"ΠΑΝΙ MICROFIBER (GnP)",unit_price:0,unit:"ΤΜΧ",category:"ΑΝΑΛΩΣΙΜΑ"},
  {row:153,code:null,name:"ΣΑΚΟΥΛΕΣ ΜΑΥΡΕΣ 110Χ130 (GnP)",unit_price:0,unit:"ΚΙΛΟ",category:"ΑΝΑΛΩΣΙΜΑ"},
  {row:154,code:null,name:"ΣΑΚΟΥΛAKIA ΓΡΑΦΕΙΟΥ WC 50*50(50ΤΜΧ) (GnP)",unit_price:0,unit:"ΤΜΧ",category:"ΑΝΑΛΩΣΙΜΑ"},
  {row:155,code:null,name:"ΣΑΚΟΥΛΕΣ ΜΑΥΡΕΣ 70Χ90 (GnP)",unit_price:0,unit:"ΚΙΛΟ",category:"ΑΝΑΛΩΣΙΜΑ"},
  {row:156,code:null,name:"ΣΑΚΟΥΛΕΣ ΜΑΥΡΕΣ 70Χ90 ΡΟΛΟ 10ΤΜΧ (GnP)",unit_price:0,unit:"ΤΜΧ",category:"ΑΝΑΛΩΣΙΜΑ"},
  {row:157,code:null,name:"ΣΦΟΥΓΓΑΡΙ ΠΡΑΣ/ΚΙΤΡΙΝΟ 6ΤΜΧ (GnP)",unit_price:0,unit:"ΤΜΧ",category:"ΑΝΑΛΩΣΙΜΑ"},
  {row:158,code:null,name:"ΣΥΡΜΑ ΑΤΣΑΛΙ 40GR (GnP)",unit_price:0,unit:"ΤΜΧ",category:"ΑΝΑΛΩΣΙΜΑ"},
  {row:159,code:null,name:"ΧΕΙΡΟΠΕΤΣΕΤΕΣ AUTOCUT WC (GnP) (6TMX KOYTA)",unit_price:0,unit:"ΤΜΧ",category:"ΑΝΑΛΩΣΙΜΑ"},
  {row:160,code:null,name:"ΧΑΡΤΙ ΚΟΥΖΙΝΑΣ ΒΙΟΜΗΧ. 4ΚΙΛΑ(2ΤΜΧ/ΣΥΣΚ.) (GNP)",unit_price:0,unit:"ΤΜΧ",category:"ΑΝΑΛΩΣΙΜΑ"},
  {row:161,code:null,name:"ΧΑΡΤΙ ΥΓΕΙΑΣ ΦΥΛΛΟ ΦΥΛΛΟ (GnP)",unit_price:0,unit:"ΤΜΧ",category:"ΑΝΑΛΩΣΙΜΑ"},
  {row:162,code:null,name:"CAFIZA ΒΑΡΕΛΑΚΙ ΚΑΘΑΡΙΣΜΟΥ",unit_price:0,unit:"ΤΜΧ",category:"ΑΝΑΛΩΣΙΜΑ"},
  // ΤΣΑΪ
  {row:166,code:"3028",name:"ΜΑΥΡΟ ΤΣΑΙ (BRAND CONNECT) BLACK VELVET",unit_price:0,unit:"ΤΜΧ",category:"ΤΣΑΪ"},
  {row:167,code:"3022",name:"ΧΑΜΟΜΗΛΙ (BRAND CONNECT)",unit_price:0,unit:"ΤΜΧ",category:"ΤΣΑΪ"},
  {row:168,code:"3025",name:"ΛΟΥΙΖΑ (BRAND CONNECT)",unit_price:0,unit:"ΤΜΧ",category:"ΤΣΑΪ"},
  {row:169,code:"3023",name:"ΧΑΛΑΡΩΣΗΣ GENTLE EVENING (BRAND CONNECT)",unit_price:0,unit:"ΤΜΧ",category:"ΤΣΑΪ"},
  {row:170,code:"3027",name:"ΠΡΑΣΙΝΟ ΤΣΑΙ RITUAL GREEN (BRAND CONNECT)",unit_price:0,unit:"ΤΜΧ",category:"ΤΣΑΪ"},
  {row:171,code:"40836",name:"ENERGY MATCHA (400ΓΡ)",unit_price:0,unit:"ΣΑΚΟΥΛΑ",category:"ΤΣΑΪ"},
  {row:172,code:"40833-0.40",name:"DETOX MATCHA (400ΓΡ)",unit_price:0,unit:"ΣΑΚΟΥΛΑ",category:"ΤΣΑΪ"},
  {row:173,code:"40815-0.40",name:"PEACH (400ΓΡ)",unit_price:0,unit:"ΣΑΚΟΥΛΑ",category:"ΤΣΑΪ"},
  // ΣΥΣΚΕΥΑΣΙΕΣ
  {row:177,code:null,name:"ΑΝΑΔΕΥΤΗΡΕΣ ΞΥΛΙΝΟΙ ΚΟΥΜΠ. (κ.ΚΩΣΤΑΣ)",unit_price:0,unit:"ΣΑΚΟΥΛΑ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:178,code:null,name:"ΑΝΑΔΕΥΤΗΡΕΣ ΞΥΛΙΝΟΙ ΜΕ ΠΥΡΙΤΥΛΙΓΜΑ (GnP)",unit_price:0,unit:"ΣΑΚΟΥΛΑ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:179,code:"G 05016",name:"ΚΑΛΑΜΑΚΙ FREDDO ΧΑΡΤΙΝΑ 1000ΤΜΧ (GnP)",unit_price:0,unit:"ΣΑΚΟΥΛΑ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:180,code:null,name:"ΚΑΛΑΜΑΚΙ JUMBO ΧΑΡΤΙΝΑ 500ΤΜΧ (κ.ΚΩΣΤΑΣ)",unit_price:0,unit:"ΣΑΚΟΥΛΑ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:181,code:"G 04946",name:"ΚΑΛΑΜΑΚΙ ΧΥΜΟΥ 1000ΤΜΧ (GnP)",unit_price:0,unit:"ΣΑΚΟΥΛΑ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:182,code:"5031",name:"ΚΑΠΑΚΙ ΧΑΡΤΙΝΟ 4oz (COSTA COFFEE) 100ΤΜΧ",unit_price:0,unit:"ΔΕΣΜΙΔΑ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:183,code:"631",name:"ΚΑΠΑΚΙ ΠΙΠΙΛΑ ΖΑΧ/ΜΟ 14OZ 50ΑΔΑ (G&P)",unit_price:0,unit:"ΔΕΣΜΙΔΑ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:184,code:"G 04251",name:"ΚΑΠΑΚΙ ΠΙΠΙΛΑ ΖΑΧ/ΜΟ 8ΟΖ 50ΑΔΑ (G&P)",unit_price:0,unit:"ΔΕΣΜΙΔΑ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:185,code:"657",name:"ΚΑΠΑΚΙ MEALS 50ΑΔΑ (κ.ΚΩΣΤΑΣ)",unit_price:0,unit:"ΔΕΣΜΙΔΑ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:186,code:"G 04518",name:"ΚΑΠΑΚΙ PET ΙΣΙΟ 14oz (G&P) 100ΑΔΑ",unit_price:0,unit:"ΔΕΣΜΙΔΑ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:187,code:"G 04546",name:"ΚΑΠΑΚΙ PET ΙΣΙΟ ΧΥΜΟΥ (G&P) 100AΔΑ",unit_price:0,unit:"ΔΕΣΜΙΔΑ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:188,code:"220",name:"ΚΡΑΦΤ ΘΗΚΗ ΤΟΣΤ+ΚΑΠΑΚΙ (25ΤΜΧ) (κ.ΚΩΣΤΑΣ)",unit_price:0,unit:"ΤΜΧ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:189,code:"301",name:"ΚΟΥΤΑΛΙ ΞΥΛΙΝΟ 100ΤΜΧ (κ.ΚΩΣΤΑΣ)",unit_price:0,unit:"ΣΑΚΟΥΛΑ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:190,code:"414",name:"ΜΑΧΑΙΡΙ ΞΥΛΙΝΟ 100ΤΜΧ (κ.ΚΩΣΤΑΣ)",unit_price:0,unit:"ΣΑΚΟΥΛΑ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:191,code:"83",name:"ΠΙΡΟΥΝΙ ΞΥΛΙΝΟ 100ΤΜΧ (κ.ΚΩΣΤΑΣ)",unit_price:0,unit:"ΣΑΚΟΥΛΑ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:192,code:"G 03689",name:"ΠΟΤΗΡΙ 350ML (G&P) 50ΑΔΑ 504",unit_price:0,unit:"ΔΕΣΜΙΔΑ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:193,code:"G 04206",name:"ΠΟΤΗΡΙ PET 105 U 450ml 50ΑΔΑ (G&P)",unit_price:0,unit:"ΔΕΣΜΙΔΑ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:194,code:null,name:"ΠΟΤΗΡΙ REUSABLE CUP",unit_price:0,unit:"ΤΜΧ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:195,code:null,name:"ΠΟΤΗΡΙ ΔΙΑΦΑΝΟ ΝΕΡΟΥ(G&P) 50ΑΔΑ",unit_price:0,unit:"ΔΕΣΜΙΔΑ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:196,code:"G 04932",name:"ΠΟΤΗΡΙ ΧΑΡΤΙΝΟ 14ΟΖ(G&P) 25ΑΔΑ",unit_price:0,unit:"ΔΕΣΜΙΔΑ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:197,code:"G 05015",name:"ΠΟΤΗΡΙ ΧΑΡΤΙΝΟ 8ΟΖ(G&P) 25ΑΔΑ",unit_price:0,unit:"ΔΕΣΜΙΔΑ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:198,code:"5001",name:"ΠΟΤΗΡΙ ΧΑΡΤΙΝΟ 4oz (COSTA COFFEE) 100ΤΜΧ",unit_price:0,unit:"ΔΕΣΜΙΔΑ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:199,code:"G 03876",name:'ΣΚΕΥΟΣ "DELI" (en.bowls) + KΑΠΑΚΙ (G&P)',unit_price:0,unit:"TMX",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:202,code:"644",name:"ΣΚΕΥΟΣ MEALS 50ΑΔΑ (κ.ΚΩΣΤΑΣ)",unit_price:0,unit:"ΔΕΣΜΙΔΑ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:203,code:"G 00098",name:"ΣΩΣΑΚΙ ΔΙΑΦΑΝΟ ΜΕ ΚΑΠΑΚΙ 50ΑΔΑ (G&P)",unit_price:0,unit:"ΔΕΣΜΙΔΑ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:204,code:"7091",name:"FILTER BELOGIA CONE 4 (COSTA COFFEE)",unit_price:0,unit:"TMX",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:205,code:"G 01913",name:"ΧΑΡΤΙ VEGETAL ΛΕΥΚΟ ΤΥΛΙΓΜΑΤΟΣ (GnP)",unit_price:0,unit:"ΚΙΛΟ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:206,code:null,name:"ΧΑΡΤΙΝΗ ΘΗΚΗ 2ΠΛΗ (GnP)",unit_price:0,unit:"ΤΜΧ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:207,code:null,name:"ΧΑΡΤΙΝΗ ΘΗΚΗ 4ΠΛΗ (GnP)",unit_price:0,unit:"ΤΜΧ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:208,code:null,name:"ΧΑΡΤΙΝΗ ΣΑΚΟΥΛΑ JOIN ΜΕ ΧΕΡΟΥΛΙ (GnP)",unit_price:0,unit:"ΤΜΧ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:209,code:null,name:"ΧΑΡΤΙΝΗ ΣΑΚΟΥΛΑ ΤΥΠΩΜΕΝΗ JOIN ΠΟΡΤΟΚΑΛΙ (GnP)",unit_price:0,unit:"ΤΜΧ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:210,code:"G 04152",name:"ΧΑΡΤOΣΑΚΟΥΛΑ ΣΑΝΤΟΥΙΤΣ ΠΟΡΤΟΚΑΛΙ JOIN (GnP)",unit_price:0,unit:"ΚΙΛΟ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:211,code:null,name:"ΧΑΡΤΟΠΕΤΣΕΤΑ (667ΤΜΧ) (ΒΕΡΓΙΝΑ)",unit_price:0,unit:"ΣΑΚΟΥΛΑ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:212,code:null,name:"ΣΑΚΟΥΛΑΚΙ GLASSBAG ΜΗΧΑΝΗΣ (GnP)",unit_price:0,unit:"ΤΜΧ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
  {row:213,code:null,name:"ΣΑΚΟΥΛΑΚΙ ΔΙΑΦΑΝΟ (GnP)",unit_price:0,unit:"ΚΙΛΟ",category:"ΣΥΣΚΕΥΑΣΙΕΣ"},
];

// Build a lookup map for fast access
const ITEM_MAP = {};
ITEMS.forEach(item => { ITEM_MAP[item.row] = item; });

// ==============================
// EXCEL SPECIAL ROWS (for export)
// ==============================
const SPECIAL_ROWS = {
  1:   {data:["ΚΩΔΙΚΟΣ","FOOD","TIMH MONΑΔΑΣ","ΠΟΣΟΤΗΤΑ","ΣΥΝΟΛΟ","ΜΟΝΑΔΑ ΜΕΤΡΗΣΗΣ"]},
  77:  {data:[null,"ΣΥΝΟΛΟ FOOD",0,null,null,null], formula:"SUM(E2:E73)"},
  78:  {empty:true},
  79:  {data:[null,"COFFEE",0,null,"ΣΥΝΟΛΟ","ΜΟΝΑΔΑ ΜΕΤΡΗΣΗΣ"]},
  97:  {data:[null,"ΣΥΝΟΛΟ COFFEE",0,null,null,null], formula:"SUM(E80:E96)"},
  98:  {empty:true},
  99:  {data:[null,"JUICES",0,null,"ΣΥΝΟΛΟ","ΜΟΝΑΔΑ ΜΕΤΡΗΣΗΣ"]},
  137: {data:[null,"ΣΥΝΟΛΟ JUICES",0,null,null,null], formula:"SUM(E100:E135)"},
  138: {empty:true},
  139: {data:[null,"ΑΝΑΛΩΣΙΜΑ",0,null,"ΣΥΝΟΛΟ","ΜΟΝΑΔΑ ΜΕΤΡΗΣΗΣ"]},
  163: {data:[null,"ΣΥΝΟΛΟ ",0,null,null,null], formula:"SUM(E140:E162)"},
  164: {empty:true},
  165: {data:[null,"ΤΣΑΪ",0,null,"ΣΥΝΟΛΟ","ΜΟΝΑΔΑ ΜΕΤΡΗΣΗΣ"]},
  174: {data:[null,"ΣΥΝΟΛΟ ΤΣΑΪ",0,null,null,null], formula:"SUM(E166:E173)"},
  175: {empty:true},
  176: {data:[null,"ΣΥΣΚΕΥΑΣΙΕΣ",0,null,"ΣΥΝΟΛΟ","ΜΟΝΑΔΑ ΜΕΤΡΗΣΗΣ"]},
  200: {empty:true},
  201: {empty:true},
  214: {data:[null,"ΣΥΝΟΛΟ ΣΥΣΚΕΥΑΣΙΑΣ",null,null,null,null], formula:"SUM(E177:E213)"},
};

// ==============================
// CONSTANTS
// ==============================
const SUPABASE_URL_KEY = 'join_sb_url';
const SUPABASE_KEY_KEY = 'join_sb_key';
const STORE_KEY        = 'join_store_id';
const DATE_KEY         = 'join_count_date';
/** Same key as tameioV2 (Ταμείο) — shared preference if both apps run on same origin */
const DARK_MODE_KEY   = 'darkMode';
const LEGACY_THEME_KEY = 'join_theme';
const PIN_KEY          = 'join_admin_pin';
const DEFAULT_PIN      = '1234';
const CATEGORIES       = ['FOOD','COFFEE','JUICES','ΑΝΑΛΩΣΙΜΑ','ΤΣΑΪ','ΣΥΣΚΕΥΑΣΙΕΣ'];
const CAT_LABELS       = {'FOOD':'🥗 FOOD','COFFEE':'☕ COFFEE','JUICES':'🥤 JUICES','ΑΝΑΛΩΣΙΜΑ':'🧻 ΑΝΑΛΩΣΙΜΑ','ΤΣΑΪ':'🍵 ΤΣΑΪ','ΣΥΣΚΕΥΑΣΙΕΣ':'📦 ΣΥΣΚΕΥΑΣΙΕΣ'};
const STORE_NAMES      = {1:'Cosmos', 2:'Πατρών', 3:'Λευκός', 4:'Ποσειδώνιο', 5:'OneSalica'};
// Items with distinct per-slot sub-locations and independent tare weights.
// Each entry fixes the slot count and labels; tare is determined by whether each slot has a cw.
const SLOT_LABELS = {
  60: ['Μπροστά', 'Πίσω', 'ΚΤΨ'],                                         // ΣΟΛΟΜΟΣ
  45: ['Καρύδια smoothies', 'Καρύδια Φαγητό', 'Καρύδια πίσω'],            // ΚΑΡΥΔΙΑ
  29: ['Αμύγδαλα smoothies', 'Αμύγδαλα Φαγητό', 'Αμύγδαλα πίσω'],       // ΑΜΥΓΔΑΛΟ ΩΜΟ
};
// Items with named group sections where some groups have configurable θέσεις.
// configurable=true  → slot count = itemConfig.num_inputs, renders θέσεις selector
// configurable=false → always 1 slot at reservedSlot (a high index that won't collide)
const SLOT_GROUPS = {
  133: [                                             // ΦΡΑΟΥΛΑ ΚΤΨ
    { label: 'Μπροστά', configurable: true  },
    { label: 'Πίσω',    configurable: false, reservedSlot: 9 },
  ],
};
// Multi-section cards: one visual card groups logically related items.
// primary row = provides the card header name; sections listed in display order.
const ITEM_GROUPS = {
  32: {  // ΒΡΩΜΗ 1ΚΙΛ (ΤΣΑΚΙΡΗΣ)
    sections: [
      { row: 19, label: 'Μπροστά' },  // ΚΙΛΟ — configurable θέσεις + αποβαρο
      { row: 32, label: 'ΤΜΧ' },      // piece count
    ]
  }
};
// Secondary rows rendered only via their group card, never as standalone items.
const ITEM_GROUP_SECONDARY_ROWS = new Set([19]);
// Maps every row in a group (including the primary itself) → primary row.
const ITEM_GROUP_BY_ROW = {};
Object.entries(ITEM_GROUPS).forEach(([primary, group]) => {
  const p = Number(primary);
  group.sections.forEach(s => { ITEM_GROUP_BY_ROW[s.row] = p; });
});
// Default store codes (used until an admin sets custom ones in Supabase)
const DEFAULT_STORE_CODES = {1:'1111', 2:'2222', 3:'3333', 4:'4444', 5:'5555'};
const DEFAULT_MASTER_CODE = '0000';

// ==============================
// STATE
// ==============================
let sb             = null;
let storeId        = null;
let countDate      = null;
let counts           = {};   // { row: netQuantity }
let containerWeights = {};  // { row: weight_kg }
let itemConfig       = {};  // { [item_row]: { num_inputs, tare_count } }
let slotValues       = {};  // { [item_row]: [netVal0, netVal1, ...] }  — all stored as NET
let storeCodes       = {};  // { 'store_1': hash, ..., 'master': hash }
let activeCategory = 'FOOD';
let searchQuery    = '';
let channel        = null;
let saveTimers     = {};
let adminDirty          = false;
let adminOriginalValues = {}; // { data-row: string } captured when admin screen opens
let skippedItems   = new Set(); // item rows marked as "Δεν μετριέται"

// Keep navigation in this tab; the last selected store is not the current screen.
const VIEW_KEY = 'join_view';
let currentScreen = 'stores';
function rememberView(screen = currentScreen) {
  currentScreen = screen;
  try {
    sessionStorage.setItem(VIEW_KEY, JSON.stringify({
      screen, storeId, countDate, activeCategory, searchQuery
    }));
  } catch (_) { /* Navigation still works when browser storage is unavailable. */ }
}

function readView() {
  try { return JSON.parse(sessionStorage.getItem(VIEW_KEY)) || {}; }
  catch (_) { return {}; }
}

// ==============================
// SLOT ENCODING HELPERS
// ==============================
function slotRow(itemRow, slot) { return slot === 0 ? itemRow : itemRow + slot * 10000; }
const SKIP_MARKER_BASE = 9000; // skip state stored as item_row + 9000 in inventory_counts
function parseSlotRow(stored) {
  if (stored < 1000) return { itemRow: stored, slot: 0 };
  return { itemRow: stored % 10000, slot: Math.floor(stored / 10000) };
}

function getItemCfg(itemRow) {
  const labels = SLOT_LABELS[itemRow];
  if (labels) {
    return { num_inputs: labels.length, tare_count: labels.length };
  }
  const cfg = itemConfig[itemRow];
  const hasCw = !!containerWeights[itemRow];
  return {
    num_inputs: cfg ? cfg.num_inputs : 1,
    tare_count: cfg ? cfg.tare_count : (hasCw ? 1 : 0)
  };
}

// Returns the container weight for a specific slot.
// SLOT_LABELS / SLOT_GROUPS items use per-slot encoded keys; regular items share one key.
function getSlotCw(itemRow, slot) {
  if (SLOT_LABELS[itemRow] || SLOT_GROUPS[itemRow]) {
    return containerWeights[slotRow(itemRow, slot)];
  }
  return containerWeights[itemRow];
}

// Returns the flat list of slot indices that contribute to an item's total.
function getEffectiveSlotIndices(itemRow) {
  const groups = SLOT_GROUPS[itemRow];
  if (groups) {
    const { num_inputs } = getItemCfg(itemRow);
    return groups.flatMap(g =>
      g.configurable ? Array.from({ length: num_inputs }, (_, i) => i) : [g.reservedSlot]
    );
  }
  const { num_inputs } = getItemCfg(itemRow);
  return Array.from({ length: num_inputs }, (_, i) => i);
}

// Computes total net kg and hasAny for any ΚΙΛΟ item (respects SLOT_GROUPS reserved slots).
function computeKiloTotal(itemRow, allSlots) {
  const indices = getEffectiveSlotIndices(itemRow);
  const total  = indices.reduce((s, idx) => s + (allSlots[idx] || 0), 0);
  const hasAny = indices.some(idx => allSlots[idx] !== undefined);
  return { total, hasAny };
}

// ==============================
// UTILITIES
// ==============================
async function sha256(str) {
  const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(str + 'join-salt'));
  return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2,'0')).join('');
}

async function verifyPin(pin) {
  const stored = localStorage.getItem(PIN_KEY);
  if (!stored) return pin === DEFAULT_PIN;
  return (await sha256(pin)) === stored;
}

async function savePin(pin) {
  localStorage.setItem(PIN_KEY, await sha256(pin));
}

async function loadStoreCodes() {
  try {
    const { data } = await sb.from('store_codes').select('*');
    storeCodes = {};
    (data || []).forEach(r => { storeCodes[r.key] = r.code_hash; });
  } catch(e) {
    storeCodes = {};
  }
}

async function verifyStoreCode(n, code) {
  const hash = await sha256(code);
  const storeKey = `store_${n}`;
  // Check store-specific code (custom or default)
  const storeMatch = storeCodes[storeKey]
    ? hash === storeCodes[storeKey]
    : code === DEFAULT_STORE_CODES[n];
  if (storeMatch) return true;
  // Check master code
  const masterMatch = storeCodes['master']
    ? hash === storeCodes['master']
    : code === DEFAULT_MASTER_CODE;
  return masterMatch;
}

function todayISO() {
  return new Date().toISOString().split('T')[0];
}

function formatDate(d) {
  const [y, m, day] = d.split('-');
  return `${day}/${m}/${y}`;
}

async function toggleSkip(itemRow) {
  const isSkipped = !skippedItems.has(itemRow);
  if (isSkipped) skippedItems.add(itemRow); else skippedItems.delete(itemRow);
  const rowEl = document.getElementById(`item-row-${itemRow}`);
  if (rowEl) rowEl.classList.toggle('skipped', isSkipped);
  const btn = document.getElementById(`skip-btn-${itemRow}`);
  if (btn) {
    btn.classList.toggle('active', isSkipped);
    btn.textContent = isSkipped ? '✓ Δεν μετριέται' : 'Δεν μετριέται';
  }
  refreshTabBadges();
  await upsertCount(itemRow + SKIP_MARKER_BASE, isSkipped ? 1 : 0);
}

async function toggleGroupSkip(primaryRow) {
  const group = ITEM_GROUPS[primaryRow];
  const isSkipped = !group.sections.some(s => skippedItems.has(s.row));
  for (const s of group.sections) {
    if (isSkipped) skippedItems.add(s.row); else skippedItems.delete(s.row);
  }
  const groupEl = document.getElementById(`item-group-${primaryRow}`);
  if (groupEl) groupEl.classList.toggle('skipped', isSkipped);
  const btn = document.getElementById(`skip-btn-group-${primaryRow}`);
  if (btn) {
    btn.classList.toggle('active', isSkipped);
    btn.textContent = isSkipped ? '✓ Δεν μετριέται' : 'Δεν μετριέται';
  }
  refreshTabBadges();
  for (const s of group.sections) {
    await upsertCount(s.row + SKIP_MARKER_BASE, isSkipped ? 1 : 0);
  }
}

function h(str) {
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function showToast(msg, type = 'info') {
  document.querySelectorAll('.toast').forEach(t => t.remove());
  const el = document.createElement('div');
  el.className = `toast toast-${type}`;
  el.textContent = msg;
  document.body.appendChild(el);
  requestAnimationFrame(() => { requestAnimationFrame(() => el.classList.add('show')); });
  setTimeout(() => {
    el.classList.remove('show');
    setTimeout(() => el.remove(), 300);
  }, 2800);
}

function migrateLegacyThemeKey() {
  const dm = localStorage.getItem(DARK_MODE_KEY);
  if (dm === 'true' || dm === 'false') return;
  const legacy = localStorage.getItem(LEGACY_THEME_KEY);
  if (legacy === 'dark') {
    localStorage.setItem(DARK_MODE_KEY, 'true');
    localStorage.removeItem(LEGACY_THEME_KEY);
  } else if (legacy === 'light') {
    localStorage.setItem(DARK_MODE_KEY, 'false');
    localStorage.removeItem(LEGACY_THEME_KEY);
  }
}

function isDarkMode() {
  return document.body.classList.contains('dark-mode');
}

/** Matches tameioV2: only localStorage, default light */
function getDarkModePreference() {
  migrateLegacyThemeKey();
  return localStorage.getItem(DARK_MODE_KEY) === 'true';
}

function applyDarkMode(on) {
  document.body.classList.toggle('dark-mode', !!on);
  const meta = document.querySelector('meta[name="theme-color"]');
  if (meta) meta.content = on ? '#000000' : '#F0F4C3';
}

function initThemeFromStorage() {
  applyDarkMode(getDarkModePreference());
}

function syncThemeToggleButtons() {
  const dark = isDarkMode();
  const title = dark ? 'Φωτεινό θέμα' : 'Σκοτεινό θέμα';
  document.querySelectorAll('.theme-toggle').forEach(btn => {
    btn.title = title;
    btn.setAttribute('aria-label', title);
    btn.textContent = dark ? '☀️' : '🌙';
  });
}

function toggleTheme() {
  const next = !isDarkMode();
  localStorage.setItem(DARK_MODE_KEY, next ? 'true' : 'false');
  applyDarkMode(next);
  syncThemeToggleButtons();
}

function themeToggleHtml() {
  const dark = isDarkMode();
  const title = dark ? 'Φωτεινό θέμα' : 'Σκοτεινό θέμα';
  return `<button type="button" class="icon-btn theme-toggle" onclick="toggleTheme()" title="${h(title)}" aria-label="${h(title)}">${dark ? '☀️' : '🌙'}</button>`;
}

function isKilo(item) {
  return item.unit.trim().toUpperCase() === 'ΚΙΛΟ';
}

function fmtQty(item, val) {
  if (isKilo(item)) return val.toFixed(3);
  return val % 1 === 0 ? String(val) : parseFloat(val.toFixed(3)).toString();
}

function isGroupFilled(primaryRow) {
  return ITEM_GROUPS[primaryRow].sections.some(s => counts[s.row] !== undefined);
}

function getItemRowEl(itemRow) {
  return document.getElementById(`item-row-${itemRow}`)
      || document.getElementById(`item-group-${ITEM_GROUP_BY_ROW[itemRow]}`);
}

/** ΝΕΡΑ AΥΡΑ 500ML — sold in packs of 24; slot0 = # of full packs, slot1 = loose ΤΜΧ */
const ITEM_ROW_AURA_WATER_500 = 86;
const AURA_WATER_PACK_UNIT = 24;

function isAuraWaterPackRow(itemRow) {
  return itemRow === ITEM_ROW_AURA_WATER_500;
}

function auraWaterPackTotal(slots) {
  const p = slots[0];
  const l = slots[1];
  const packs = p !== undefined ? Math.max(0, Math.floor(Number(p))) : 0;
  const loose = l !== undefined ? Math.max(0, Math.floor(Number(l))) : 0;
  return packs * AURA_WATER_PACK_UNIT + loose;
}

function auraWaterPackHasAnyInput(slots) {
  return slots && (slots[0] !== undefined || slots[1] !== undefined);
}

// ==============================
// SUPABASE
// ==============================
function initSupabase() {
  const url = localStorage.getItem(SUPABASE_URL_KEY);
  const key = localStorage.getItem(SUPABASE_KEY_KEY);
  if (!url || !key) return false;
  sb = supabase.createClient(url, key);
  return true;
}

async function loadData() {
  try {
    const [wRes, cRes, cfgRes] = await Promise.all([
      sb.from('container_weights').select('*').eq('store_id', storeId),
      sb.from('inventory_counts').select('*').eq('store_id', storeId).eq('count_date', countDate),
      sb.from('item_config').select('*').eq('store_id', storeId)
    ]);
    if (wRes.error) throw wRes.error;
    if (cRes.error) throw cRes.error;
    // cfgRes error is non-fatal (table may not exist yet)
    if (cfgRes.error) console.warn('item_config load skipped:', cfgRes.error.message);

    containerWeights = {};
    (wRes.data || []).forEach(r => { containerWeights[r.item_row] = parseFloat(r.weight_kg); });

    itemConfig = {};
    (cfgRes.data || []).forEach(r => {
      itemConfig[r.item_row] = { num_inputs: r.num_inputs, tare_count: r.tare_count };
    });

    slotValues = {};
    skippedItems = new Set();
    (cRes.data || []).forEach(r => {
      if (r.quantity === null) return;
      if (r.item_row > SKIP_MARKER_BASE) {
        const realRow = r.item_row - SKIP_MARKER_BASE;
        if (ITEM_MAP[realRow] && r.quantity > 0) skippedItems.add(realRow);
        return;
      }
      const { itemRow, slot } = parseSlotRow(r.item_row);
      if (!slotValues[itemRow]) slotValues[itemRow] = [];
      slotValues[itemRow][slot] = parseFloat(r.quantity);
    });

    // Legacy Aura water: one row stored total ΤΜΧ in slot 0 only → split to 0 packs + loose (same total)
    const aw = ITEM_ROW_AURA_WATER_500;
    const awSlots = slotValues[aw];
    if (awSlots && awSlots[1] === undefined && awSlots[0] !== undefined) {
      const loose = Math.max(0, Math.floor(parseFloat(awSlots[0])));
      if (loose === 0) {
        delete slotValues[aw];
        await deleteCount(aw);
      } else {
        slotValues[aw][0] = 0;
        slotValues[aw][1] = loose;
        await upsertCount(aw, 0);
        await upsertCount(slotRow(aw, 1), loose);
      }
    }

    counts = {};
    for (const [rowStr, slots] of Object.entries(slotValues)) {
      const itemRow = parseInt(rowStr);
      if (isAuraWaterPackRow(itemRow)) {
        if (!auraWaterPackHasAnyInput(slots)) continue;
        const total = auraWaterPackTotal(slots);
        if (total > 0) counts[itemRow] = total;
        continue;
      }
      const { total, hasAny } = computeKiloTotal(itemRow, slots);
      if (hasAny) counts[itemRow] = total;
    }

    return true;
  } catch(e) {
    console.error('Load error:', e);
    return false;
  }
}

async function upsertCount(itemRow, netQty) {
  try {
    const { error } = await sb.from('inventory_counts').upsert(
      { store_id: storeId, count_date: countDate, item_row: itemRow, quantity: netQty, updated_at: new Date().toISOString() },
      { onConflict: 'store_id,count_date,item_row' }
    );
    if (error) throw error;
  } catch(e) {
    console.error('Save error:', e);
    showToast('Σφάλμα αποθήκευσης', 'error');
  }
}

async function deleteCount(itemRow) {
  try {
    await sb.from('inventory_counts').delete()
      .eq('store_id', storeId).eq('count_date', countDate).eq('item_row', itemRow);
  } catch(e) {
    console.error('Delete error:', e);
  }
}

function subscribeToChanges() {
  if (channel) sb.removeChannel(channel);
  channel = sb.channel(`store-${storeId}-${countDate}`)
    .on('postgres_changes', {
      event: '*', schema: 'public', table: 'inventory_counts',
      filter: `store_id=eq.${storeId}`
    }, payload => {
      const rec = payload.new || {};
      const storedRow = rec.item_row;
      const qty = (rec.quantity !== null && rec.quantity !== undefined) ? parseFloat(rec.quantity) : null;

      // Skip marker: sync "Δεν μετριέται" state from another device
      if (storedRow > SKIP_MARKER_BASE) {
        const realRow = storedRow - SKIP_MARKER_BASE;
        const item = ITEM_MAP[realRow];
        if (!item || rec.count_date !== countDate) return;
        const isSkipped = qty !== null && qty > 0;
        if (isSkipped) skippedItems.add(realRow); else skippedItems.delete(realRow);
        const groupPrimary = ITEM_GROUP_BY_ROW[realRow];
        if (groupPrimary !== undefined) {
          const groupSkipped = ITEM_GROUPS[groupPrimary].sections.some(s => skippedItems.has(s.row));
          const groupEl = document.getElementById(`item-group-${groupPrimary}`);
          if (groupEl) groupEl.classList.toggle('skipped', groupSkipped);
          const btn = document.getElementById(`skip-btn-group-${groupPrimary}`);
          if (btn) {
            btn.classList.toggle('active', groupSkipped);
            btn.textContent = groupSkipped ? '✓ Δεν μετριέται' : 'Δεν μετριέται';
          }
        } else {
          const rowEl = document.getElementById(`item-row-${realRow}`);
          if (rowEl) rowEl.classList.toggle('skipped', isSkipped);
          const btn = document.getElementById(`skip-btn-${realRow}`);
          if (btn) {
            btn.classList.toggle('active', isSkipped);
            btn.textContent = isSkipped ? '✓ Δεν μετριέται' : 'Δεν μετριέται';
          }
        }
        refreshTabBadges();
        return;
      }

      const { itemRow, slot } = parseSlotRow(storedRow);
      const item = ITEM_MAP[itemRow];
      if (!item || rec.count_date !== countDate) return;

      // Update slotValues
      if (!slotValues[itemRow]) slotValues[itemRow] = [];
      if (qty !== null) slotValues[itemRow][slot] = qty;
      else delete slotValues[itemRow][slot];

      // Recompute total
      const allSlots = slotValues[itemRow] || [];
      let total, hasAny;
      if (isAuraWaterPackRow(itemRow)) {
        hasAny = auraWaterPackHasAnyInput(allSlots);
        total = auraWaterPackTotal(allSlots);
        if (!hasAny || total <= 0) delete counts[itemRow];
        else counts[itemRow] = total;
      } else {
        ({ total, hasAny } = computeKiloTotal(itemRow, allSlots));
        if (hasAny) counts[itemRow] = total; else delete counts[itemRow];
      }

      // Update input if not focused
      const input = document.querySelector(`[data-row="${itemRow}"][data-slot="${slot}"]`);
      if (input && document.activeElement !== input) {
        const slotCw = getSlotCw(itemRow, slot);
        const isTare = isKilo(item) && !!slotCw;
        if (isTare) {
          input.value = qty !== null ? (qty + slotCw).toFixed(3) : '';
          const netEl = document.getElementById(`net-${itemRow}-${slot}`);
          if (netEl) {
            netEl.classList.toggle('empty', qty === null);
            netEl.innerHTML = qty !== null
              ? `<span class="net-value">${qty.toFixed(3)}</span><span class="net-unit">kg</span>`
              : `<span class="net-value placeholder">—</span>`;
          }
        } else if (isAuraWaterPackRow(itemRow)) {
          input.value = qty !== null ? String(Math.max(0, Math.floor(qty))) : '';
        } else {
          input.value = qty !== null ? qty.toFixed(3) : '';
        }
        if (!isAuraWaterPackRow(itemRow)) {
          const totalEl = document.getElementById(`total-${itemRow}`);
          if (totalEl) totalEl.textContent = hasAny ? fmtQty(item, total) : '—';
          const rowEl = getItemRowEl(itemRow);
          const gp = ITEM_GROUP_BY_ROW[itemRow];
          if (rowEl) rowEl.classList.toggle('filled', gp !== undefined ? isGroupFilled(gp) : hasAny);
        }
      }
      if (isAuraWaterPackRow(itemRow)) {
        const totalEl = document.getElementById(`total-${itemRow}`);
        if (totalEl) totalEl.textContent = hasAny && total > 0 ? String(total) : '—';
        const rowEl = document.getElementById(`item-row-${itemRow}`);
        if (rowEl) rowEl.classList.toggle('filled', hasAny && total > 0);
      }
      refreshTabBadges();
    })
    .subscribe(status => {
      const dot = document.getElementById('conn-dot');
      if (dot) {
        dot.className = `conn-dot ${status === 'SUBSCRIBED' ? 'connected' : 'disconnected'}`;
        dot.title = status === 'SUBSCRIBED' ? 'Συνδεδεμένο' : 'Αποσυνδεδεμένο';
      }
    });
}

// ==============================
// ITEM HELPERS
// ==============================
function getInputDisplayValue(item) {
  const net = counts[item.row];
  if (net === undefined) return '';
  const cw = containerWeights[item.row];
  if (isKilo(item) && cw) return (net + cw).toFixed(3);
  return net % 1 === 0 ? String(net) : parseFloat(net.toFixed(3)).toString();
}

function processInput(item, raw) {
  if (raw === '' || raw === null) return null;
  const val = parseFloat(raw);
  if (isNaN(val) || val < 0) return null;
  const cw = containerWeights[item.row];
  if (isKilo(item) && cw) return Math.max(0, val - cw);
  return val;
}

// ==============================
// SCREEN: SETUP
// ==============================
function renderSetupScreen() {
  rememberView('setup');
  document.getElementById('app').innerHTML = `
    <div class="screen setup-screen">
      <div class="theme-corner">${themeToggleHtml()}</div>
      <div class="setup-card">
        <div class="logo">
          <div class="logo-icon">J</div>
          <h1>JOIN Απογραφή</h1>
          <p class="subtitle">Σύνδεση Supabase</p>
        </div>
        <div class="form-group">
          <label>Supabase Project URL</label>
          <input type="url" id="sb-url" placeholder="https://xxxx.supabase.co"
            value="${h(localStorage.getItem(SUPABASE_URL_KEY) || '')}">
        </div>
        <div class="form-group">
          <label>Anon Key</label>
          <input type="text" id="sb-key" placeholder="eyJhbGci..."
            value="${h(localStorage.getItem(SUPABASE_KEY_KEY) || '')}">
        </div>
        <button class="btn-primary" onclick="saveSetup()">Αποθήκευση &amp; Συνέχεια</button>
        <div id="setup-error" class="error-msg"></div>
      </div>
    </div>`;
}

async function saveSetup() {
  const url = document.getElementById('sb-url').value.trim();
  const key = document.getElementById('sb-key').value.trim();
  const errEl = document.getElementById('setup-error');
  if (!url || !key) { errEl.textContent = 'Συμπληρώστε και τα δύο πεδία.'; return; }
  errEl.textContent = 'Σύνδεση...';
  try {
    const client = supabase.createClient(url, key);
    const { error } = await client.from('inventory_counts').select('id').limit(1);
    if (error) throw error;
    localStorage.setItem(SUPABASE_URL_KEY, url);
    localStorage.setItem(SUPABASE_KEY_KEY, key);
    sb = client;
    renderStoreScreen();
  } catch(e) {
    errEl.textContent = 'Σφάλμα σύνδεσης. Ελέγξτε URL και Anon Key.';
  }
}

// ==============================
// SCREEN: STORE SELECTION
// ==============================
function renderStoreScreen() {
  rememberView('stores');
  const saved = localStorage.getItem(DATE_KEY) || todayISO();
  document.getElementById('app').innerHTML = `
    <div class="screen store-screen">
      <div class="theme-corner">${themeToggleHtml()}</div>
      <div class="store-card">
        <div class="logo">
          <div class="logo-icon">J</div>
          <h1>JOIN Απογραφή</h1>
          <p class="subtitle">Επιλογή Καταστήματος</p>
        </div>
        <div class="form-group date-group">
          <label>Ημερομηνία Απογραφής</label>
          <input type="date" id="count-date" value="${saved}">
        </div>
        <div class="store-grid">
          ${[1,2,3,4,5].map(n => `
            <button class="store-btn" onclick="selectStore(${n})">
              <span class="store-emoji">🏪</span>
              <span class="store-label">${STORE_NAMES[n]}</span>
            </button>`).join('')}
        </div>
        <button class="btn-secondary small" onclick="renderSetupScreen()">⚙ Ρυθμίσεις Supabase</button>
      </div>
    </div>`;
}

async function selectStore(n) {
  const dateEl = document.getElementById('count-date');
  const date = dateEl ? dateEl.value : todayISO();

  // If already authenticated this session for this store, skip code entry
  if (sessionStorage.getItem(`join_auth_${n}`) === '1') {
    await enterStore(n, date);
    return;
  }

  // Load current codes then show entry modal
  await loadStoreCodes();
  showStoreCodeModal(n, date);
}

function showStoreCodeModal(n, date) {
  const modal = document.createElement('div');
  modal.className = 'modal-overlay';
  modal.id = 'store-code-modal';
  modal.innerHTML = `
    <div class="modal-card">
      <h2>${STORE_NAMES[n]}</h2>
      <p style="text-align:center;font-size:13px;color:var(--text-secondary);margin-bottom:4px">Εισάγετε κωδικό καταστήματος</p>
      <div class="pin-inputs">
        ${[0,1,2,3].map(i => `<input type="password" class="pin-digit store-code-d" id="scd${i}" maxlength="1" inputmode="numeric" pattern="[0-9]">`).join('')}
      </div>
      <div id="store-code-err" class="error-msg" style="text-align:center"></div>
      <div class="modal-actions">
        <button class="btn-secondary" onclick="document.getElementById('store-code-modal').remove()">Ακύρωση</button>
        <button class="btn-primary" onclick="submitStoreCode(${n},'${date}')">Είσοδος</button>
      </div>
    </div>`;
  document.body.appendChild(modal);
  const digits = modal.querySelectorAll('.store-code-d');
  digits.forEach((d, i) => {
    d.addEventListener('input', () => {
      d.value = d.value.replace(/\D/g,'').slice(-1);
      if (d.value && i < 3) digits[i+1].focus();
      if (d.value && i === 3) submitStoreCode(n, date);
    });
    d.addEventListener('keydown', e => {
      if (e.key === 'Backspace' && !d.value && i > 0) digits[i-1].focus();
    });
  });
  digits[0].focus();
}

async function submitStoreCode(n, date) {
  const digits = document.querySelectorAll('.store-code-d');
  const code = Array.from(digits).map(d => d.value).join('');
  if (code.length < 4) {
    document.getElementById('store-code-err').textContent = 'Εισάγετε 4 ψηφία';
    return;
  }
  const ok = await verifyStoreCode(n, code);
  if (ok) {
    sessionStorage.setItem(`join_auth_${n}`, '1');
    const hash = await sha256(code);
    const isMaster = storeCodes['master'] ? hash === storeCodes['master'] : code === DEFAULT_MASTER_CODE;
    sessionStorage.setItem(`join_auth_type_${n}`, isMaster ? 'master' : 'store');
    document.getElementById('store-code-modal').remove();
    await enterStore(n, date);
  } else {
    document.getElementById('store-code-err').textContent = 'Λάθος κωδικός';
    digits.forEach(d => { d.value = ''; });
    digits[0].focus();
  }
}

async function enterStore(n, date) {
  storeId = n;
  countDate = date;
  // skippedItems is now loaded from Supabase inside loadData()
  localStorage.setItem(STORE_KEY, n);
  localStorage.setItem(DATE_KEY, date);
  document.getElementById('app').innerHTML = `
    <div class="screen loading-screen">
      <div class="spinner"></div>
      <p>Φόρτωση δεδομένων...</p>
    </div>`;
  const ok = await loadData();
  if (ok) { renderCountingScreen(); subscribeToChanges(); }
  else { showToast('Σφάλμα φόρτωσης', 'error'); renderStoreScreen(); }
}

// ==============================
// SCREEN: COUNTING
// ==============================
function renderCountingScreen() {
  rememberView('counting');
  document.getElementById('app').innerHTML = `
    <div class="screen counting-screen">
      <header class="app-header">
        <div class="header-left">
          <button class="back-btn" onclick="goBack()" title="Αλλαγή">←</button>
          <div class="header-title">
            <strong>${STORE_NAMES[storeId]}</strong>
            <span class="header-date">${formatDate(countDate)}</span>
          </div>
        </div>
        <div class="header-right">
          ${themeToggleHtml()}
          <span id="conn-dot" class="conn-dot disconnected" title="Αποσυνδεδεμένο"></span>
          <button class="icon-btn" onclick="openAdminPin()" title="Ρυθμίσεις διαχειριστή">⚙</button>
        </div>
      </header>

      <div class="search-bar-wrap">
        <input type="search" id="search-input" class="search-input"
          placeholder="🔍 Αναζήτηση είδους..."
          oninput="onSearch(this.value)"
          value="${h(searchQuery)}">
      </div>

      <div class="tabs-wrap" id="tabs-wrap"${searchQuery ? ' style="display:none"' : ''}>${buildTabs()}</div>

      <div class="items-container" id="items-container">${buildItemsList()}</div>

      <div class="bottom-bar">
        <span class="total-info" id="total-info">${totalCountedText()}</span>
        <button class="btn-export" id="export-btn" onclick="exportToExcel()">📥 Εξαγωγή Excel</button>
      </div>
    </div>`;

  // Scroll active tab into view
  const activeTab = document.querySelector('.tab-btn.active');
  if (activeTab) activeTab.scrollIntoView({behavior:'instant', block:'nearest', inline:'center'});
}

function buildTabs() {
  return CATEGORIES.map(cat => {
    const items = ITEMS.filter(i => i.category === cat);
    const n = items.filter(i => counts[i.row] !== undefined || skippedItems.has(i.row)).length;
    const complete = n === items.length;
    return `<button class="tab-btn ${cat === activeCategory ? 'active' : ''}"
      onclick="switchTab('${cat}')">${CAT_LABELS[cat]}
      <span class="tab-badge ${n > 0 ? 'has-counts' : ''}${complete ? ' complete' : ''}">${n}/${items.length}</span>
    </button>`;
  }).join('');
}

function buildItemsList() {
  let list;
  if (searchQuery) {
    const q = searchQuery.toLowerCase();
    const matched = ITEMS.filter(i =>
      i.name.toLowerCase().includes(q) || (i.code && i.code.toLowerCase().includes(q))
    );
    // Deduplicate: secondary rows resolve to their group primary
    const seenPrimary = new Set();
    list = [];
    for (const item of matched) {
      const primary = ITEM_GROUP_BY_ROW[item.row];
      const keyRow = primary !== undefined ? primary : item.row;
      if (seenPrimary.has(keyRow)) continue;
      seenPrimary.add(keyRow);
      list.push(ITEM_MAP[keyRow] || item);
    }
  } else {
    list = ITEMS.filter(i => i.category === activeCategory);
  }

  if (!list.length) return '<div class="empty-state">Δεν βρέθηκαν είδη</div>';

  let html = '';
  let lastCat = null;
  for (const item of list) {
    if (ITEM_GROUP_SECONDARY_ROWS.has(item.row)) continue;
    if (searchQuery && item.category !== lastCat) {
      html += `<div class="category-divider">${CAT_LABELS[item.category] || item.category}</div>`;
      lastCat = item.category;
    }
    if (ITEM_GROUPS[item.row]) {
      html += buildGroupRow(item.row);
    } else {
      html += buildItemRow(item);
    }
  }
  return html;
}

function buildGroupRow(primaryRow) {
  const primary = ITEM_MAP[primaryRow];
  const group = ITEM_GROUPS[primaryRow];
  const isSkipped = group.sections.some(s => skippedItems.has(s.row));
  const filled = isGroupFilled(primaryRow);
  const step = '0.001';
  const mode = 'decimal';

  const sectionsHtml = group.sections.map(section => {
    const sItem = ITEM_MAP[section.row];
    if (!sItem) return '';

    if (isKilo(sItem)) {
      const cfg = getItemCfg(section.row);
      const { num_inputs, tare_count } = cfg;
      const slots = slotValues[section.row] || [];

      const numOpts = Array.from({length:10},(_,i)=>
        `<option value="${i+1}"${num_inputs===i+1?' selected':''}>${i+1}</option>`).join('');
      const tareOpts = Array.from({length:num_inputs+1},(_,i)=>
        `<option value="${i}"${tare_count===i?' selected':''}>${i}</option>`).join('');
      const cfgHtml = `
        <div class="item-cfg-selects">
          <div class="cfg-sel-wrap">
            <span class="cfg-label">Θέσεις</span>
            <select class="cfg-select" onchange="onItemCfgChange(${section.row},'num_inputs',this.value)">${numOpts}</select>
          </div>
          <div class="cfg-sel-wrap">
            <span class="cfg-label">Αποβαρο</span>
            <select class="cfg-select" id="tare-sel-${section.row}" onchange="onItemCfgChange(${section.row},'tare_count',this.value)">${tareOpts}</select>
          </div>
        </div>`;

      const slotsHtml = Array.from({ length: num_inputs }, (_, slot) => {
        const slotCw = getSlotCw(section.row, slot);
        const isTare = !!slotCw;
        const netVal = slots[slot];
        const dispVal = isTare && netVal !== undefined ? (netVal + slotCw).toFixed(3) : (netVal !== undefined ? netVal.toFixed(3) : '');
        const slotLabel = num_inputs > 1 ? `<span class="slot-label">Θ${slot + 1}</span>` : '';
        const netDisplay = isTare ? `
          <div class="net-display ${netVal !== undefined ? '' : 'empty'}" id="net-${section.row}-${slot}">
            ${netVal !== undefined ? `<span class="net-value">${netVal.toFixed(3)}</span><span class="net-unit">kg</span>` : `<span class="net-value placeholder">—</span>`}
          </div>` : '';
        return `
          <div class="item-slot">
            ${slotLabel}
            <div class="item-input-wrap slot-input-wrap">
              <div class="input-group">
                <input type="number" class="qty-input${isTare ? ' gross-input' : ''}"
                  data-row="${section.row}" data-slot="${slot}"
                  placeholder="${isTare ? 'Μεικτό' : '0'}"
                  value="${dispVal}" min="0" step="${step}" inputmode="${mode}"
                  oninput="onSlotInput(${section.row}, ${slot}, this.value)">
                <span class="unit-label">kg</span>
              </div>
              ${netDisplay}
            </div>
          </div>`;
      }).join('');

      const totalKilo = counts[section.row];
      const totalBar = num_inputs > 1 ? `
        <div class="item-total-bar">
          <span>📊 Σύνολο:</span>
          <span class="total-value" id="total-${section.row}">${totalKilo !== undefined ? totalKilo.toFixed(3) : '—'}</span>
          <span>kg</span>
        </div>` : '';

      return `
        <div class="slot-group">
          <div class="slot-group-header">
            <span class="slot-group-label">${h(section.label)}</span>
            ${cfgHtml}
          </div>
          ${slotsHtml}
          ${totalBar}
        </div>`;
    } else {
      // Non-ΚΙΛΟ section: simple piece count
      const net = counts[section.row];
      const displayVal = net !== undefined ? (net % 1 === 0 ? String(net) : parseFloat(net.toFixed(3)).toString()) : '';
      return `
        <div class="slot-group">
          <div class="slot-group-header">
            <span class="slot-group-label">${h(section.label)}</span>
          </div>
          <div class="item-slot">
            <div class="item-input-wrap slot-input-wrap">
              <div class="input-group">
                <input type="number" class="qty-input" data-row="${section.row}" data-slot="0"
                  placeholder="0" value="${displayVal}" min="0" step="1" inputmode="numeric"
                  oninput="onSlotInput(${section.row}, 0, this.value)">
                <span class="unit-label">${h(sItem.unit)}</span>
              </div>
            </div>
          </div>
        </div>`;
    }
  }).join('');

  return `
    <div class="item-group item-row${filled ? ' filled' : ''}${isSkipped ? ' skipped' : ''}" id="item-group-${primaryRow}">
      <div class="item-header-row">
        <div class="item-info">
          <span class="item-name">${h(primary.name)}</span>
          ${primary.code ? `<span class="item-code">${h(primary.code)}</span>` : ''}
        </div>
      </div>
      <div class="item-slots">
        ${sectionsHtml}
      </div>
      <div class="skip-row">
        <button id="skip-btn-group-${primaryRow}" class="skip-btn${isSkipped ? ' active' : ''}" onclick="toggleGroupSkip(${primaryRow})">${isSkipped ? '✓ Δεν μετριέται' : 'Δεν μετριέται'}</button>
      </div>
    </div>`;
}

function buildItemRow(item) {
  const cfg = getItemCfg(item.row);
  const cw = containerWeights[item.row];
  const total = counts[item.row];
  const filled = total !== undefined && total > 0;
  const step = '0.001';
  const mode = 'decimal';

  if (!isKilo(item) && isAuraWaterPackRow(item.row)) {
    const slots = slotValues[item.row] || [];
    const packsVal = slots[0] !== undefined ? String(Math.max(0, Math.floor(slots[0]))) : '';
    const looseVal = slots[1] !== undefined ? String(Math.max(0, Math.floor(slots[1]))) : '';
    const tot = auraWaterPackTotal(slots);
    const filled = tot > 0;
    return `
      <div class="item-row multi-slot aura-pack-row${filled ? ' filled' : ''}${skippedItems.has(item.row) ? ' skipped' : ''}" id="item-row-${item.row}">
        <div class="item-header-row">
          <div class="item-info">
            <span class="item-name">${h(item.name)}</span>
            ${item.code ? `<span class="item-code">${h(item.code)}</span>` : ''}
          </div>
        </div>
        <div class="item-slots">
          <div class="item-slot">
            <span class="slot-label">Κιβώτια ×${AURA_WATER_PACK_UNIT}</span>
            <div class="item-input-wrap slot-input-wrap">
              <div class="input-group">
                <input type="number" class="qty-input" data-row="${item.row}" data-slot="0"
                  placeholder="0" value="${packsVal}" min="0" step="1" inputmode="numeric"
                  oninput="onSlotInput(${item.row}, 0, this.value)">
                <span class="unit-label">κιβ.</span>
              </div>
            </div>
          </div>
          <div class="item-slot">
            <span class="slot-label">Εκτός κιβωτίου</span>
            <div class="item-input-wrap slot-input-wrap">
              <div class="input-group">
                <input type="number" class="qty-input" data-row="${item.row}" data-slot="1"
                  placeholder="0" value="${looseVal}" min="0" step="1" inputmode="numeric"
                  oninput="onSlotInput(${item.row}, 1, this.value)">
                <span class="unit-label">ΤΜΧ</span>
              </div>
            </div>
          </div>
          <div class="item-total-bar">
            <span>📊 Σύνολο:</span>
            <span class="total-value" id="total-${item.row}">${tot > 0 ? String(tot) : '—'}</span>
            <span>ΤΜΧ</span>
          </div>
        </div>
        <div class="skip-row">
          <button id="skip-btn-${item.row}" class="skip-btn${skippedItems.has(item.row) ? ' active' : ''}" onclick="toggleSkip(${item.row})">${skippedItems.has(item.row) ? '✓ Δεν μετριέται' : 'Δεν μετριέται'}</button>
        </div>
      </div>`;
  }

  if (!isKilo(item)) {
    const { num_inputs } = cfg;
    const slots = slotValues[item.row] || [];
    const total = counts[item.row];
    const filled = total !== undefined && total > 0;

    const numOpts = Array.from({length:10},(_,i)=>
      `<option value="${i+1}"${num_inputs===i+1?' selected':''}>${i+1}</option>`).join('');
    const cfgHtml = `
      <div class="item-cfg-selects">
        <div class="cfg-sel-wrap">
          <span class="cfg-label">Θέσεις</span>
          <select class="cfg-select" onchange="onItemCfgChange(${item.row},'num_inputs',this.value)">${numOpts}</select>
        </div>
      </div>`;

    const slotsHtml = Array.from({ length: num_inputs }, (_, slot) => {
      const val = slots[slot];
      const displayVal = val !== undefined ? fmtQty(item, val) : '';
      const slotLabel = num_inputs > 1 ? `<span class="slot-label">Θ${slot + 1}</span>` : '';
      return `
        <div class="item-slot">
          ${slotLabel}
          <div class="item-input-wrap slot-input-wrap">
            <div class="input-group">
              <input type="number" class="qty-input" data-row="${item.row}" data-slot="${slot}"
                placeholder="0" value="${displayVal}" min="0" step="1" inputmode="numeric"
                oninput="onSlotInput(${item.row}, ${slot}, this.value)">
              <span class="unit-label">${h(item.unit)}</span>
            </div>
          </div>
        </div>`;
    }).join('');

    const totalBar = num_inputs > 1 ? `
      <div class="item-total-bar">
        <span>📊 Σύνολο:</span>
        <span class="total-value" id="total-${item.row}">${total !== undefined ? fmtQty(item, total) : '—'}</span>
        <span>${h(item.unit)}</span>
      </div>` : '';

    return `
      <div class="item-row${filled ? ' filled' : ''}${num_inputs > 1 ? ' multi-slot' : ''}${skippedItems.has(item.row) ? ' skipped' : ''}" id="item-row-${item.row}">
        <div class="item-header-row">
          <div class="item-info">
            <span class="item-name">${h(item.name)}</span>
            ${item.code ? `<span class="item-code">${h(item.code)}</span>` : ''}
          </div>
          ${cfgHtml}
        </div>
        <div class="item-slots">
          ${slotsHtml}
          ${totalBar}
        </div>
        <div class="skip-row">
          <button id="skip-btn-${item.row}" class="skip-btn${skippedItems.has(item.row) ? ' active' : ''}" onclick="toggleSkip(${item.row})">${skippedItems.has(item.row) ? '✓ Δεν μετριέται' : 'Δεν μετριέται'}</button>
        </div>
      </div>`;
  }

  // SLOT_GROUPS item — named sections, one configurable (θέσεις), one fixed
  if (SLOT_GROUPS[item.row]) {
    const groups = SLOT_GROUPS[item.row];
    const { num_inputs } = cfg;
    const slots = slotValues[item.row] || [];
    const displayTotal = counts[item.row];

    const groupsHtml = groups.map(group => {
      const groupSlots = group.configurable
        ? Array.from({ length: num_inputs }, (_, i) => i)
        : [group.reservedSlot];

      const cfgHtml = group.configurable ? (() => {
        const numOpts = Array.from({length:10},(_,i)=>
          `<option value="${i+1}"${num_inputs===i+1?' selected':''}>${i+1}</option>`).join('');
        return `<div class="item-cfg-selects">
          <div class="cfg-sel-wrap">
            <span class="cfg-label">Θέσεις</span>
            <select class="cfg-select" onchange="onItemCfgChange(${item.row},'num_inputs',this.value)">${numOpts}</select>
          </div>
        </div>`;
      })() : '';

      const slotsHtml = groupSlots.map(slot => {
        const slotCw = getSlotCw(item.row, slot);
        const isTare = !!slotCw;
        const netVal = slots[slot];
        const dispVal = isTare && netVal !== undefined ? (netVal + slotCw).toFixed(3) : (netVal !== undefined ? netVal.toFixed(3) : '');
        const slotLabel = (group.configurable && num_inputs > 1) ? `<span class="slot-label">Θ${slot + 1}</span>` : '';
        const netDisplay = isTare ? `
          <div class="net-display ${netVal !== undefined ? '' : 'empty'}" id="net-${item.row}-${slot}">
            ${netVal !== undefined ? `<span class="net-value">${netVal.toFixed(3)}</span><span class="net-unit">kg</span>` : `<span class="net-value placeholder">—</span>`}
          </div>` : '';
        return `
          <div class="item-slot">
            ${slotLabel}
            <div class="item-input-wrap slot-input-wrap">
              <div class="input-group">
                <input type="number" class="qty-input${isTare ? ' gross-input' : ''}"
                  data-row="${item.row}" data-slot="${slot}"
                  placeholder="${isTare ? 'Μεικτό' : '0'}"
                  value="${dispVal}" min="0" step="${step}" inputmode="${mode}"
                  oninput="onSlotInput(${item.row}, ${slot}, this.value)">
                <span class="unit-label">kg</span>
              </div>
              ${netDisplay}
            </div>
          </div>`;
      }).join('');

      return `
        <div class="slot-group">
          <div class="slot-group-header">
            <span class="slot-group-label">${group.label}</span>
            ${cfgHtml}
          </div>
          ${slotsHtml}
        </div>`;
    }).join('');

    return `
      <div class="item-row${filled ? ' filled' : ''} multi-slot${skippedItems.has(item.row) ? ' skipped' : ''}" id="item-row-${item.row}">
        <div class="item-header-row">
          <div class="item-info">
            <span class="item-name">${h(item.name)}</span>
            ${item.code ? `<span class="item-code">${h(item.code)}</span>` : ''}
          </div>
        </div>
        <div class="item-slots">
          ${groupsHtml}
          <div class="item-total-bar">
            <span>📊 Σύνολο:</span>
            <span class="total-value" id="total-${item.row}">${displayTotal !== undefined ? displayTotal.toFixed(3) : '—'}</span>
            <span>kg</span>
          </div>
        </div>
        <div class="skip-row">
          <button id="skip-btn-${item.row}" class="skip-btn${skippedItems.has(item.row) ? ' active' : ''}" onclick="toggleSkip(${item.row})">${skippedItems.has(item.row) ? '✓ Δεν μετριέται' : 'Δεν μετριέται'}</button>
        </div>
      </div>`;
  }

  // ΚΙΛΟ item — may have multiple slots
  const { num_inputs, tare_count } = cfg;
  const slots = slotValues[item.row] || [];
  const hasSlotLabels = !!SLOT_LABELS[item.row];

  const slotsHtml = Array.from({ length: num_inputs }, (_, slot) => {
    const slotCw = getSlotCw(item.row, slot);
    const isTare = !!slotCw;
    const netVal = slots[slot];  // stored as net
    const dispVal = isTare && netVal !== undefined ? (netVal + slotCw).toFixed(3) : (netVal !== undefined ? netVal.toFixed(3) : '');
    const labelText = hasSlotLabels
      ? (SLOT_LABELS[item.row][slot] || `Θ${slot + 1}`)
      : (num_inputs > 1 ? `Θ${slot + 1}` : '');
    const slotLabel = labelText ? `<span class="slot-label">${labelText}</span>` : '';

    const netDisplay = isTare ? `
      <div class="net-display ${netVal !== undefined ? '' : 'empty'}" id="net-${item.row}-${slot}">
        ${netVal !== undefined
          ? `<span class="net-value">${netVal.toFixed(3)}</span><span class="net-unit">kg</span>`
          : `<span class="net-value placeholder">—</span>`}
      </div>` : '';

    return `
      <div class="item-slot">
        ${slotLabel}
        <div class="item-input-wrap slot-input-wrap">
          <div class="input-group">
            <input type="number" class="qty-input${isTare ? ' gross-input' : ''}"
              data-row="${item.row}" data-slot="${slot}"
              placeholder="${isTare ? 'Μεικτό' : '0'}"
              value="${dispVal}"
              min="0" step="${step}" inputmode="${mode}"
              oninput="onSlotInput(${item.row}, ${slot}, this.value)">
            <span class="unit-label">kg</span>
          </div>
          ${netDisplay}
        </div>
      </div>`;
  }).join('');

  const totalBar = num_inputs > 1 ? `
    <div class="item-total-bar">
      <span>📊 Σύνολο:</span>
      <span class="total-value" id="total-${item.row}">${total !== undefined ? total.toFixed(3) : '—'}</span>
      <span>kg</span>
    </div>` : '';

  const cfgSelectsHtml = hasSlotLabels ? '' : (() => {
    const numOpts = Array.from({length:10},(_,i)=>
      `<option value="${i+1}"${num_inputs===i+1?' selected':''}>${i+1}</option>`).join('');
    const tareOpts = Array.from({length:num_inputs+1},(_,i)=>
      `<option value="${i}"${tare_count===i?' selected':''}>${i}</option>`).join('');
    return `
      <div class="item-cfg-selects">
        <div class="cfg-sel-wrap">
          <span class="cfg-label">Θέσεις</span>
          <select class="cfg-select" onchange="onItemCfgChange(${item.row},'num_inputs',this.value)">${numOpts}</select>
        </div>
        <div class="cfg-sel-wrap">
          <span class="cfg-label">Αποβαρο</span>
          <select class="cfg-select" id="tare-sel-${item.row}" onchange="onItemCfgChange(${item.row},'tare_count',this.value)">${tareOpts}</select>
        </div>
      </div>`;
  })();

  return `
    <div class="item-row${filled ? ' filled' : ''}${num_inputs > 1 ? ' multi-slot' : ''}${skippedItems.has(item.row) ? ' skipped' : ''}" id="item-row-${item.row}">
      <div class="item-header-row">
        <div class="item-info">
          <span class="item-name">${h(item.name)}</span>
          ${item.code ? `<span class="item-code">${h(item.code)}</span>` : ''}
        </div>
        ${cfgSelectsHtml}
      </div>
      <div class="item-slots">
        ${slotsHtml}
        ${totalBar}
      </div>
      <div class="skip-row">
        <button id="skip-btn-${item.row}" class="skip-btn${skippedItems.has(item.row) ? ' active' : ''}" onclick="toggleSkip(${item.row})">${skippedItems.has(item.row) ? '✓ Δεν μετριέται' : 'Δεν μετριέται'}</button>
      </div>
    </div>`;
}

function switchTab(cat) {
  activeCategory = cat;
  searchQuery = '';
  rememberView();
  const si = document.getElementById('search-input');
  if (si) si.value = '';
  const tw = document.getElementById('tabs-wrap');
  if (tw) tw.innerHTML = buildTabs();
  const ic = document.getElementById('items-container');
  if (ic) ic.innerHTML = buildItemsList();
  const tabsWrap = document.getElementById('tabs-wrap');
  if (tabsWrap) tabsWrap.style.display = '';
  const activeTab = document.querySelector('.tab-btn.active');
  if (activeTab) activeTab.scrollIntoView({behavior:'smooth', block:'nearest', inline:'center'});
}

function onSearch(q) {
  searchQuery = q;
  rememberView();
  const ic = document.getElementById('items-container');
  if (ic) ic.innerHTML = buildItemsList();
  const tw = document.getElementById('tabs-wrap');
  if (tw) tw.style.display = q ? 'none' : '';
}

function totalCountedText() {
  const n = ITEMS.filter(i => counts[i.row] !== undefined || skippedItems.has(i.row)).length;
  return `${n} / ${ITEMS.length} μετρήθηκαν`;
}

function refreshTabBadges() {
  const tw = document.getElementById('tabs-wrap');
  if (tw && !searchQuery) tw.innerHTML = buildTabs();
  const ti = document.getElementById('total-info');
  if (ti) ti.textContent = totalCountedText();
}

// ==============================
// ITEM INPUT
// ==============================
function onSlotInput(itemRow, slot, raw) {
  if (typeof raw === 'string') raw = raw.replace(',', '.');
  const item = ITEM_MAP[itemRow];
  if (!item) return;

  if (isAuraWaterPackRow(itemRow)) {
    let v = null;
    if (raw !== '' && raw !== null) {
      const n = parseInt(raw, 10);
      if (!isNaN(n) && n >= 0) v = n;
    }
    if (!slotValues[itemRow]) slotValues[itemRow] = [];
    if (v !== null) slotValues[itemRow][slot] = v;
    else delete slotValues[itemRow][slot];

    const allSlots = slotValues[itemRow] || [];
    const hasEntry = auraWaterPackHasAnyInput(allSlots);
    const total = auraWaterPackTotal(allSlots);
    if (!hasEntry || total <= 0) delete counts[itemRow];
    else counts[itemRow] = total;

    const totalEl = document.getElementById(`total-${itemRow}`);
    if (totalEl) totalEl.textContent = hasEntry && total > 0 ? String(total) : '—';
    const rowEl = document.getElementById(`item-row-${itemRow}`);
    if (rowEl) rowEl.classList.toggle('filled', total > 0 && hasEntry);

    refreshTabBadges();

    clearTimeout(saveTimers[`aura-${itemRow}`]);
    saveTimers[`aura-${itemRow}`] = setTimeout(() => {
      const sl = slotValues[itemRow] || [];
      const has = auraWaterPackHasAnyInput(sl);
      const tot = auraWaterPackTotal(sl);
      if (!has || tot <= 0) {
        deleteCount(itemRow);
        deleteCount(slotRow(itemRow, 1));
      } else {
        if (sl[0] !== undefined) upsertCount(itemRow, Math.floor(sl[0]));
        else deleteCount(itemRow);
        if (sl[1] !== undefined) upsertCount(slotRow(itemRow, 1), Math.floor(sl[1]));
        else deleteCount(slotRow(itemRow, 1));
      }
    }, 600);
    return;
  }

  const cw = getSlotCw(itemRow, slot);
  const isTare = isKilo(item) && !!cw;

  let netVal = null;
  if (raw !== '' && raw !== null) {
    const v = parseFloat(raw);
    if (!isNaN(v)) {
      netVal = isTare ? Math.max(0, v - cw) : v;
    }
  }

  // Update net display for tare slots
  if (isTare) {
    const netEl = document.getElementById(`net-${itemRow}-${slot}`);
    if (netEl) {
      if (netVal !== null) {
        netEl.classList.remove('empty');
        netEl.innerHTML = `<span class="net-value">${netVal.toFixed(3)}</span><span class="net-unit">kg</span>`;
      } else {
        netEl.classList.add('empty');
        netEl.innerHTML = `<span class="net-value placeholder">—</span>`;
      }
    }
  }

  // Update slotValues
  if (!slotValues[itemRow]) slotValues[itemRow] = [];
  if (netVal !== null) {
    slotValues[itemRow][slot] = netVal;
  } else {
    delete slotValues[itemRow][slot];
  }

  // Recompute total
  const allSlots = slotValues[itemRow] || [];
  const { total, hasAny } = computeKiloTotal(itemRow, allSlots);

  if (hasAny) {
    counts[itemRow] = total;
  } else {
    delete counts[itemRow];
  }

  // Update total display
  const totalEl = document.getElementById(`total-${itemRow}`);
  if (totalEl) totalEl.textContent = hasAny ? fmtQty(item, total) : '—';

  // Update filled state
  const rowEl = getItemRowEl(itemRow);
  const gp = ITEM_GROUP_BY_ROW[itemRow];
  if (rowEl) rowEl.classList.toggle('filled', gp !== undefined ? isGroupFilled(gp) : hasAny);

  refreshTabBadges();

  // Debounce save
  const key = `${itemRow}-${slot}`;
  clearTimeout(saveTimers[key]);
  saveTimers[key] = setTimeout(() => {
    const sr = slotRow(itemRow, slot);
    if (netVal !== null) upsertCount(sr, netVal);
    else deleteCount(sr);
  }, 600);
}

// Backward-compatibility alias
function onItemInput(itemRow, raw) { onSlotInput(itemRow, 0, raw); }

function onItemCfgChange(itemRow, type, newVal) {
  const v = parseInt(newVal);
  const existing = itemConfig[itemRow] || { num_inputs: 1, tare_count: containerWeights[itemRow] ? 1 : 0 };
  const cfg = { ...existing };
  const cw = containerWeights[itemRow] || 0;

  if (type === 'num_inputs') {
    cfg.num_inputs = v;
    cfg.tare_count = Math.min(cfg.tare_count, v);
  } else {
    const oldTare = cfg.tare_count;
    const newTare = v;
    // Reinterpret values for slots whose tare status changed
    if (cw > 0 && slotValues[itemRow]) {
      const lo = Math.min(oldTare, newTare);
      const hi = Math.max(oldTare, newTare);
      for (let slot = lo; slot < hi; slot++) {
        const stored = slotValues[itemRow][slot];
        if (stored === undefined) continue;
        // plain→tare: input showed net, now treated as gross → new net = stored − cw
        // tare→plain: input showed gross (stored+cw), now that becomes the plain value → new net = stored + cw
        const newNet = newTare > oldTare
          ? Math.max(0, stored - cw)
          : stored + cw;
        slotValues[itemRow][slot] = newNet;
        const sr = slotRow(itemRow, slot);
        clearTimeout(saveTimers[`${itemRow}-${slot}`]);
        saveTimers[`${itemRow}-${slot}`] = setTimeout(() => upsertCount(sr, newNet), 800);
      }
    }
    cfg.tare_count = newTare;
    // Recompute total
    const allSlots = slotValues[itemRow] || [];
    const total = allSlots.slice(0, cfg.num_inputs).reduce((s, val) => s + (val || 0), 0);
    const hasAny = allSlots.slice(0, cfg.num_inputs).some(val => val !== undefined);
    if (hasAny) counts[itemRow] = total; else delete counts[itemRow];
  }

  itemConfig[itemRow] = cfg;

  const groupPrimary = ITEM_GROUP_BY_ROW[itemRow];
  if (groupPrimary !== undefined) {
    const groupEl = document.getElementById(`item-group-${groupPrimary}`);
    if (groupEl) {
      const tmp = document.createElement('div');
      tmp.innerHTML = buildGroupRow(groupPrimary);
      groupEl.replaceWith(tmp.firstElementChild);
    }
  } else {
    const rowEl = document.getElementById(`item-row-${itemRow}`);
    if (rowEl) {
      const tmp = document.createElement('div');
      tmp.innerHTML = buildItemRow(ITEM_MAP[itemRow]);
      rowEl.replaceWith(tmp.firstElementChild);
    }
  }

  clearTimeout(saveTimers[`cfg-${itemRow}`]);
  saveTimers[`cfg-${itemRow}`] = setTimeout(() => saveItemCfg(itemRow), 800);
}


async function saveItemCfg(itemRow) {
  const cfg = itemConfig[itemRow];
  if (!cfg) return;
  try {
    await sb.from('item_config').upsert(
      { store_id: storeId, item_row: itemRow, num_inputs: cfg.num_inputs, tare_count: cfg.tare_count, updated_at: new Date().toISOString() },
      { onConflict: 'store_id,item_row' }
    );
  } catch(e) {
    console.error('item_config save error:', e);
  }
}

// ==============================
// BACK NAVIGATION
// ==============================
function goBack() {
  if (channel) { sb.removeChannel(channel); channel = null; }
  renderStoreScreen();
}

// ==============================
// ADMIN — PIN MODAL
// ==============================
function openAdminPin() {
  const modal = document.createElement('div');
  modal.className = 'modal-overlay';
  modal.id = 'pin-modal';
  const isFirstTime = !localStorage.getItem(PIN_KEY);
  modal.innerHTML = `
    <div class="modal-card">
      <h2>Κωδικός Διαχειριστή</h2>
      ${isFirstTime ? '<p class="info-msg">Πρώτη φορά; Χρησιμοποιήστε τον κωδικό <strong>1234</strong></p>' : ''}
      <div class="pin-inputs" id="pin-inputs">
        ${[0,1,2,3].map(i => `<input type="password" class="pin-digit" id="pd${i}" maxlength="1" inputmode="numeric" pattern="[0-9]">`).join('')}
      </div>
      <div id="pin-error" class="error-msg" style="text-align:center"></div>
      <div class="modal-actions">
        <button class="btn-secondary" onclick="document.getElementById('pin-modal').remove()">Ακύρωση</button>
        <button class="btn-primary" onclick="submitAdminPin()">Είσοδος</button>
      </div>
    </div>`;
  document.body.appendChild(modal);
  const digits = modal.querySelectorAll('.pin-digit');
  digits.forEach((d, i) => {
    d.addEventListener('input', () => {
      d.value = d.value.replace(/\D/g,'').slice(-1);
      if (d.value && i < 3) digits[i+1].focus();
      if (d.value && i === 3) submitAdminPin();
    });
    d.addEventListener('keydown', e => {
      if (e.key === 'Backspace' && !d.value && i > 0) { digits[i-1].focus(); }
    });
  });
  digits[0].focus();
}

async function submitAdminPin() {
  const digits = document.querySelectorAll('#pin-modal .pin-digit');
  const pin = Array.from(digits).map(d => d.value).join('');
  if (pin.length < 4) {
    document.getElementById('pin-error').textContent = 'Εισάγετε 4 ψηφία';
    return;
  }
  const ok = await verifyPin(pin);
  if (ok) {
    document.getElementById('pin-modal').remove();
    renderAdminScreen();
  } else {
    document.getElementById('pin-error').textContent = 'Λάθος κωδικός';
    digits.forEach(d => { d.value = ''; });
    digits[0].focus();
  }
}

function buildCodesSection() {
  const isMaster = sessionStorage.getItem(`join_auth_type_${storeId}`) === 'master';
  const storeRows = isMaster
    ? [1,2,3,4,5].map(n => `
        <div class="admin-item-row">
          <div class="admin-item-info">
            <span class="item-name">${STORE_NAMES[n]}</span>
            <span class="item-code">Κατάστημα ${n}</span>
          </div>
          <div class="admin-input-wrap">
            <input type="password" class="store-code-input" data-key="store_${n}"
              placeholder="——" maxlength="4" inputmode="numeric" pattern="[0-9]{4}"
              style="width:64px;text-align:center">
          </div>
        </div>`).join('')
    : `
        <div class="admin-item-row">
          <div class="admin-item-info">
            <span class="item-name">${STORE_NAMES[storeId]}</span>
            <span class="item-code">Κατάστημα ${storeId}</span>
          </div>
          <div class="admin-input-wrap">
            <input type="password" class="store-code-input" data-key="store_${storeId}"
              placeholder="——" maxlength="4" inputmode="numeric" pattern="[0-9]{4}"
              style="width:64px;text-align:center">
          </div>
        </div>`;

  const masterRow = isMaster ? `
        <div class="admin-item-row">
          <div class="admin-item-info">
            <span class="item-name">Master (όλα τα καταστήματα)</span>
          </div>
          <div class="admin-input-wrap">
            <input type="password" class="store-code-input" data-key="master"
              placeholder="——" maxlength="4" inputmode="numeric" pattern="[0-9]{4}"
              style="width:64px;text-align:center">
          </div>
        </div>` : '';

  return `
    <div class="admin-section">
      <h3 class="admin-cat-title">🔐 Κωδικοί Πρόσβασης</h3>
      <p style="font-size:13px;color:var(--text-secondary);margin:0 0 12px">Εισάγετε νέο 4-ψήφιο κωδικό για να τον αλλάξετε. Αφήστε κενό για να διατηρήσετε τον τρέχοντα.</p>
      ${storeRows}
      ${masterRow}
      <div class="admin-actions" style="margin-top:12px">
        <button class="btn-primary" onclick="saveStoreCodes()">💾 Αποθήκευση Κωδικών</button>
      </div>
      <div id="codes-msg" class="success-msg" style="text-align:center"></div>
    </div>`;
}

// ==============================
// SCREEN: ADMIN
// ==============================
function adminWeightChanged(a, b) {
  const pa = parseFloat(a), pb = parseFloat(b);
  if (isNaN(pa) && isNaN(pb)) return false;
  if (isNaN(pa) || isNaN(pb)) return true;
  return pa !== pb;
}

function onAdminWeightInput(inputEl) {
  const key = inputEl.dataset.row;
  const changed = adminWeightChanged(inputEl.value, adminOriginalValues[key] ?? '');

  const rowEl = inputEl.closest('.admin-item-row');
  const warn = rowEl && rowEl.nextElementSibling;
  if (warn && warn.classList.contains('admin-field-warn')) {
    warn.style.display = changed ? 'block' : 'none';
  }

  adminDirty = Array.from(document.querySelectorAll('.admin-weight-input'))
    .some(inp => adminWeightChanged(inp.value, adminOriginalValues[inp.dataset.row] ?? ''));
}

function goBackFromAdmin() {
  if (!adminDirty) { renderCountingScreen(); return; }
  const modal = document.createElement('div');
  modal.className = 'modal-overlay';
  modal.id = 'admin-back-modal';
  modal.innerHTML = `
    <div class="modal-card">
      <h2>Μη αποθηκευμένες αλλαγές</h2>
      <p style="text-align:center;color:var(--text-secondary);margin:8px 0 16px">Έχεις αλλαγές που δεν αποθηκεύτηκαν. Θες να φύγεις χωρίς αποθήκευση;</p>
      <div class="modal-actions">
        <button class="btn-secondary" onclick="document.getElementById('admin-back-modal').remove()">Παραμονή</button>
        <button class="btn-danger" onclick="adminDirty=false;document.getElementById('admin-back-modal').remove();renderCountingScreen()">Έξοδος χωρίς αποθήκευση</button>
      </div>
    </div>`;
  document.body.appendChild(modal);
}

function renderAdminScreen() {
  rememberView('admin');
  adminDirty = false;
  const kiloItems = ITEMS.filter(isKilo);

  const sections = CATEGORIES.map(cat => {
    const items = kiloItems.filter(i => i.category === cat);
    if (!items.length) return '';
    return `
      <div class="admin-section">
        <h3 class="admin-cat-title">${CAT_LABELS[cat]}</h3>
        ${items.map(item => {
          const groups = SLOT_GROUPS[item.row];
          if (groups) {
            const frontCount = (itemConfig[item.row] && itemConfig[item.row].num_inputs) || 1;
            return groups.map(group => {
              const groupSlots = group.configurable
                ? Array.from({ length: frontCount }, (_, i) => i)
                : [group.reservedSlot];
              return groupSlots.map(slot => {
                const slotKey = slotRow(item.row, slot);
                const cwVal = containerWeights[slotKey] !== undefined ? containerWeights[slotKey] : '';
                const subLabel = (group.configurable && frontCount > 1) ? `${group.label} Θ${slot + 1}` : group.label;
                return `
                  <div class="admin-item-row">
                    <div class="admin-item-info">
                      <span class="item-name">${h(item.name)}</span>
                      <span class="item-code">${h(subLabel)}</span>
                    </div>
                    <div class="admin-input-wrap">
                      <input type="number" class="admin-weight-input"
                        data-row="${slotKey}" placeholder="—"
                        value="${cwVal}" min="0" step="0.001" inputmode="decimal"
                        oninput="onAdminWeightInput(this)">
                      <span class="unit-label">kg</span>
                    </div>
                  </div>
                  <div class="admin-field-warn" style="display:none">⚠️ Μην ξεχάσεις να κάνεις αποθήκευση!</div>`;
              }).join('');
            }).join('');
          }
          const labels = SLOT_LABELS[item.row];
          if (labels) {
            return labels.map((label, slot) => {
              const slotKey = slotRow(item.row, slot);
              const cwVal = containerWeights[slotKey] !== undefined ? containerWeights[slotKey] : '';
              return `
                <div class="admin-item-row">
                  <div class="admin-item-info">
                    <span class="item-name">${h(item.name)}</span>
                    <span class="item-code">${h(label)}</span>
                  </div>
                  <div class="admin-input-wrap">
                    <input type="number" class="admin-weight-input"
                      data-row="${slotKey}" placeholder="—"
                      value="${cwVal}" min="0" step="0.001" inputmode="decimal"
                      oninput="onAdminWeightInput(this)">
                    <span class="unit-label">kg</span>
                  </div>
                </div>
                <div class="admin-field-warn" style="display:none">⚠️ Μην ξεχάσεις να κάνεις αποθήκευση!</div>`;
            }).join('');
          }
          const cwVal = containerWeights[item.row] !== undefined ? containerWeights[item.row] : '';
          return `
            <div class="admin-item-row">
              <div class="admin-item-info">
                <span class="item-name">${h(item.name)}</span>
                ${item.code ? `<span class="item-code">${h(item.code)}</span>` : ''}
              </div>
              <div class="admin-input-wrap">
                <input type="number" class="admin-weight-input"
                  data-row="${item.row}" placeholder="—"
                  value="${cwVal}" min="0" step="0.001" inputmode="decimal"
                  oninput="onAdminWeightInput(this)">
                <span class="unit-label">kg</span>
              </div>
            </div>
            <div class="admin-field-warn" style="display:none">⚠️ Μην ξεχάσεις να κάνεις αποθήκευση!</div>`;
        }).join('')}
      </div>`;
  }).join('');

  document.getElementById('app').innerHTML = `
    <div class="screen admin-screen">
      <header class="app-header">
        <div class="header-left">
          <button class="back-btn" onclick="goBackFromAdmin()">← Πίσω</button>
          <div class="header-title"><strong>⚖️ Βάρη Δοχείων</strong></div>
        </div>
        <div class="header-right">${themeToggleHtml()}</div>
      </header>
      <div class="admin-content">
        <div class="admin-info">
          <p>Εισάγετε το βάρος του δοχείου/βάζου για κάθε είδος που μετράτε κατά βάρος. Αφήστε κενό για να μην αφαιρείται τίποτα.</p>
        </div>
        ${sections}
        <div class="admin-actions">
          <button class="btn-primary" onclick="saveContainerWeights()">💾 Αποθήκευση</button>
          <button class="btn-secondary" onclick="openChangePinModal()">🔑 Αλλαγή PIN</button>
        </div>
        <div id="admin-msg" class="success-msg" style="text-align:center"></div>

        ${buildCodesSection()}
      </div>
    </div>`;

  adminOriginalValues = {};
  document.querySelectorAll('.admin-weight-input').forEach(inp => {
    adminOriginalValues[inp.dataset.row] = inp.value;
  });
}

async function saveContainerWeights() {
  const inputs = document.querySelectorAll('.admin-weight-input');
  const upserts = [];
  const toDelete = [];

  inputs.forEach(input => {
    const row = parseInt(input.dataset.row);
    const val = parseFloat(input.value);
    if (!isNaN(val) && val > 0) {
      upserts.push({ store_id: storeId, item_row: row, weight_kg: val, updated_at: new Date().toISOString() });
    } else if (containerWeights[row] !== undefined) {
      toDelete.push(row);
    }
  });

  try {
    if (upserts.length) {
      const { error } = await sb.from('container_weights').upsert(upserts, { onConflict: 'store_id,item_row' });
      if (error) throw error;
    }
    for (const row of toDelete) {
      await sb.from('container_weights').delete().eq('store_id', storeId).eq('item_row', row);
    }
    upserts.forEach(u => { containerWeights[u.item_row] = u.weight_kg; });
    toDelete.forEach(row => { delete containerWeights[row]; });

    adminDirty = false;
    document.querySelectorAll('.admin-field-warn').forEach(w => { w.style.display = 'none'; });
    document.querySelectorAll('.admin-weight-input').forEach(inp => {
      adminOriginalValues[inp.dataset.row] = inp.value;
    });
    const msg = document.getElementById('admin-msg');
    if (msg) { msg.textContent = '✓ Αποθηκεύτηκε'; setTimeout(() => { if(msg) msg.textContent=''; }, 3000); }
    showToast('Αποθηκεύτηκε', 'success');
  } catch(e) {
    console.error(e);
    showToast('Σφάλμα αποθήκευσης', 'error');
  }
}

async function saveStoreCodes() {
  const inputs = document.querySelectorAll('.store-code-input');
  const upserts = [];

  for (const input of inputs) {
    const val = input.value.trim();
    if (!val) continue;
    if (!/^\d{4}$/.test(val)) {
      showToast('Οι κωδικοί πρέπει να είναι 4 ψηφία', 'error');
      return;
    }
    const hash = await sha256(val);
    upserts.push({ key: input.dataset.key, code_hash: hash, updated_at: new Date().toISOString() });
  }

  if (!upserts.length) {
    showToast('Δεν εισήχθη κανένας κωδικός', 'error');
    return;
  }

  try {
    const { error } = await sb.from('store_codes').upsert(upserts, { onConflict: 'key' });
    if (error) throw error;
    upserts.forEach(u => { storeCodes[u.key] = u.code_hash; });
    inputs.forEach(i => { i.value = ''; });
    const msg = document.getElementById('codes-msg');
    if (msg) { msg.textContent = '✓ Αποθηκεύτηκαν'; setTimeout(() => { if(msg) msg.textContent=''; }, 3000); }
    showToast('Κωδικοί αποθηκεύτηκαν', 'success');
  } catch(e) {
    console.error(e);
    showToast('Σφάλμα αποθήκευσης', 'error');
  }
}

// ==============================
// CHANGE PIN MODAL
// ==============================
function openChangePinModal() {
  const modal = document.createElement('div');
  modal.className = 'modal-overlay';
  modal.id = 'change-pin-modal';
  modal.innerHTML = `
    <div class="modal-card">
      <h2>Νέο PIN</h2>
      <p style="text-align:center;font-size:13px;color:var(--text-secondary);margin-bottom:4px">Εισάγετε 4-ψήφιο κωδικό</p>
      <div class="pin-inputs">
        ${[0,1,2,3].map(i => `<input type="password" class="pin-digit new-pin-d" id="npd${i}" maxlength="1" inputmode="numeric" pattern="[0-9]">`).join('')}
      </div>
      <div id="new-pin-err" class="error-msg" style="text-align:center"></div>
      <div class="modal-actions">
        <button class="btn-secondary" onclick="document.getElementById('change-pin-modal').remove()">Ακύρωση</button>
        <button class="btn-primary" onclick="submitNewPin()">Αποθήκευση</button>
      </div>
    </div>`;
  document.body.appendChild(modal);
  const digits = modal.querySelectorAll('.new-pin-d');
  digits.forEach((d, i) => {
    d.addEventListener('input', () => {
      d.value = d.value.replace(/\D/g,'').slice(-1);
      if (d.value && i < 3) digits[i+1].focus();
      if (d.value && i === 3) submitNewPin();
    });
    d.addEventListener('keydown', e => { if (e.key==='Backspace' && !d.value && i>0) digits[i-1].focus(); });
  });
  digits[0].focus();
}

async function submitNewPin() {
  const digits = document.querySelectorAll('.new-pin-d');
  const pin = Array.from(digits).map(d => d.value).join('');
  if (!/^\d{4}$/.test(pin)) {
    document.getElementById('new-pin-err').textContent = 'Απαιτούνται 4 αριθμοί';
    return;
  }
  await savePin(pin);
  document.getElementById('change-pin-modal').remove();
  showToast('PIN ενημερώθηκε', 'success');
}

// ==============================
// EXPORT TO EXCEL
// ==============================
// Base64-encoded "JOIN απογραφή κενό .xlsx" template (embedded so no server needed)
const TEMPLATE_B64 = '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';

async function exportToExcel() {
  const btn = document.getElementById('export-btn');
  if (btn) { btn.disabled = true; btn.textContent = '⏳ Εξαγωγή...'; }

  try {
    const wb = XLSX.read(TEMPLATE_B64, { type: 'base64' });
    const ws = wb.Sheets['Φύλλο1'];

    for (const rowNum of Object.keys(ITEM_MAP)) {
      const item = ITEM_MAP[rowNum];
      const qty = counts[item.row];
      if (qty !== undefined) {
        const addr = XLSX.utils.encode_cell({ r: item.row - 1, c: 3 });
        ws[addr] = { t: 'n', v: qty };
      }
    }

    const filename = `JOIN_Απογραφή_${STORE_NAMES[storeId]}_${countDate}.xlsx`;
    XLSX.writeFile(wb, filename);
    showToast('Εξαγωγή ολοκληρώθηκε', 'success');
  } catch(e) {
    console.error('Export error:', e);
    showToast('Σφάλμα εξαγωγής', 'error');
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = '📥 Εξαγωγή Excel'; }
  }
}

// ==============================
// MOBILE KEYBOARD SCROLL FIX
// ==============================
// When the soft keyboard opens the viewport shrinks and the browser's
// automatic scroll overshoots past the sticky header. We wait for the
// keyboard animation to finish, then re-scroll the input into view.
document.addEventListener('focusin', e => {
  if (!e.target.matches('.qty-input, .admin-weight-input')) return;
  const input = e.target;
  // 300ms covers both iOS and Android keyboard animation
  setTimeout(() => {
    input.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }, 300);
});

// ==============================
// INIT
// ==============================
async function init() {
  const view = readView();
  if (view.screen === 'setup') { renderSetupScreen(); return; }
  if (!initSupabase()) { renderSetupScreen(); return; }

  if (!['counting', 'admin'].includes(view.screen) ||
      ![1, 2, 3, 4, 5].includes(view.storeId) ||
      typeof view.countDate !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(view.countDate) ||
      sessionStorage.getItem(`join_auth_${view.storeId}`) !== '1') {
    renderStoreScreen();
    return;
  }

  storeId   = view.storeId;
  countDate = view.countDate;
  activeCategory = CATEGORIES.includes(view.activeCategory) ? view.activeCategory : 'FOOD';
  searchQuery = typeof view.searchQuery === 'string' ? view.searchQuery : '';

  document.getElementById('app').innerHTML = `
    <div class="screen loading-screen">
      <div class="spinner"></div>
      <p>Φόρτωση δεδομένων...</p>
    </div>`;

  const ok = await loadData();
  if (ok) {
    if (view.screen === 'admin') renderAdminScreen();
    else renderCountingScreen();
    subscribeToChanges();
  }
  else    { renderStoreScreen(); }
}

document.addEventListener('beforeinput', e => {
  if (e.data !== ',') return;
  const el = e.target;
  if (el.tagName !== 'INPUT') return;
  e.preventDefault();
  document.execCommand('insertText', false, '.');
}, true);

initThemeFromStorage();
init();
