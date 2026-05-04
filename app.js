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
const PIN_KEY          = 'join_admin_pin';
const DEFAULT_PIN      = '1234';
const CATEGORIES       = ['FOOD','COFFEE','JUICES','ΑΝΑΛΩΣΙΜΑ','ΤΣΑΪ','ΣΥΣΚΕΥΑΣΙΕΣ'];
const CAT_LABELS       = {'FOOD':'🥗 FOOD','COFFEE':'☕ COFFEE','JUICES':'🥤 JUICES','ΑΝΑΛΩΣΙΜΑ':'🧻 ΑΝΑΛΩΣΙΜΑ','ΤΣΑΪ':'🍵 ΤΣΑΪ','ΣΥΣΚΕΥΑΣΙΕΣ':'📦 ΣΥΣΚΕΥΑΣΙΕΣ'};
const STORE_NAMES      = {1:'Cosmos', 2:'Πατρών', 3:'Λευκός', 4:'Ποσειδώνιο', 5:'OneSalica'};
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

// ==============================
// SLOT ENCODING HELPERS
// ==============================
function slotRow(itemRow, slot) { return slot === 0 ? itemRow : itemRow + slot * 10000; }
function parseSlotRow(stored) {
  if (stored < 1000) return { itemRow: stored, slot: 0 };
  return { itemRow: stored % 10000, slot: Math.floor(stored / 10000) };
}

function getItemCfg(itemRow) {
  const cfg = itemConfig[itemRow];
  const hasCw = !!containerWeights[itemRow];
  return {
    num_inputs: cfg ? cfg.num_inputs : 1,
    tare_count: cfg ? cfg.tare_count : (hasCw ? 1 : 0)
  };
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

function isKilo(item) {
  return item.unit.trim().toUpperCase() === 'ΚΙΛΟ';
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
    (cRes.data || []).forEach(r => {
      if (r.quantity === null) return;
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
      const { num_inputs } = getItemCfg(itemRow);
      const total = slots.slice(0, num_inputs).reduce((s, v) => s + (v || 0), 0);
      if (total > 0) counts[itemRow] = total;
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
      const { itemRow, slot } = parseSlotRow(storedRow);
      const item = ITEM_MAP[itemRow];
      if (!item || rec.count_date !== countDate) return;

      // Update slotValues
      if (!slotValues[itemRow]) slotValues[itemRow] = [];
      if (qty !== null) slotValues[itemRow][slot] = qty;
      else delete slotValues[itemRow][slot];

      // Recompute total
      const { num_inputs, tare_count } = getItemCfg(itemRow);
      const allSlots = slotValues[itemRow] || [];
      let total, hasAny;
      if (isAuraWaterPackRow(itemRow)) {
        hasAny = auraWaterPackHasAnyInput(allSlots);
        total = auraWaterPackTotal(allSlots);
        if (!hasAny || total <= 0) delete counts[itemRow];
        else counts[itemRow] = total;
      } else {
        total = allSlots.slice(0, num_inputs).reduce((s, v) => s + (v || 0), 0);
        hasAny = allSlots.slice(0, num_inputs).some(v => v !== undefined);
        if (hasAny) counts[itemRow] = total; else delete counts[itemRow];
      }

      // Update input if not focused
      const input = document.querySelector(`[data-row="${itemRow}"][data-slot="${slot}"]`);
      if (input && document.activeElement !== input) {
        const cw = containerWeights[itemRow];
        const isTare = isKilo(item) && slot < tare_count && !!cw;
        if (isTare) {
          input.value = qty !== null ? (qty + cw).toFixed(3) : '';
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
          if (totalEl) totalEl.textContent = hasAny ? total.toFixed(3) : '—';
          const rowEl = document.getElementById(`item-row-${itemRow}`);
          if (rowEl) rowEl.classList.toggle('filled', hasAny);
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
  document.getElementById('app').innerHTML = `
    <div class="screen setup-screen">
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
  const saved = localStorage.getItem(DATE_KEY) || todayISO();
  document.getElementById('app').innerHTML = `
    <div class="screen store-screen">
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

      <div class="tabs-wrap" id="tabs-wrap">${buildTabs()}</div>

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
    const n = items.filter(i => counts[i.row] !== undefined).length;
    return `<button class="tab-btn ${cat === activeCategory ? 'active' : ''}"
      onclick="switchTab('${cat}')">${CAT_LABELS[cat]}
      <span class="tab-badge ${n > 0 ? 'has-counts' : ''}">${n}/${items.length}</span>
    </button>`;
  }).join('');
}

function buildItemsList() {
  const list = searchQuery
    ? ITEMS.filter(i => i.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
        (i.code && i.code.toLowerCase().includes(searchQuery.toLowerCase())))
    : ITEMS.filter(i => i.category === activeCategory);

  if (!list.length) return '<div class="empty-state">Δεν βρέθηκαν είδη</div>';

  let html = '';
  let lastCat = null;
  for (const item of list) {
    if (searchQuery && item.category !== lastCat) {
      html += `<div class="category-divider">${CAT_LABELS[item.category] || item.category}</div>`;
      lastCat = item.category;
    }
    html += buildItemRow(item);
  }
  return html;
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
      <div class="item-row multi-slot aura-pack-row${filled ? ' filled' : ''}" id="item-row-${item.row}">
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
      </div>`;
  }

  if (!isKilo(item)) {
    const net = counts[item.row];
    const displayVal = net !== undefined ? (net % 1 === 0 ? String(net) : parseFloat(net.toFixed(3)).toString()) : '';
    return `
      <div class="item-row ${net !== undefined ? 'filled' : ''}" id="item-row-${item.row}">
        <div class="item-header-row">
          <div class="item-info">
            <span class="item-name">${h(item.name)}</span>
            ${item.code ? `<span class="item-code">${h(item.code)}</span>` : ''}
          </div>
          <div class="item-input-wrap">
            <div class="input-group">
              <input type="number" class="qty-input" data-row="${item.row}" data-slot="0"
                placeholder="0" value="${displayVal}" min="0" step="1" inputmode="numeric"
                oninput="onSlotInput(${item.row}, 0, this.value)">
              <span class="unit-label">${h(item.unit)}</span>
            </div>
          </div>
        </div>
      </div>`;
  }

  // ΚΙΛΟ item — may have multiple slots
  const { num_inputs, tare_count } = cfg;
  const slots = slotValues[item.row] || [];

  const slotsHtml = Array.from({ length: num_inputs }, (_, slot) => {
    const isTare = slot < tare_count && !!cw;
    const netVal = slots[slot];  // stored as net
    const dispVal = isTare && netVal !== undefined ? (netVal + cw).toFixed(3) : (netVal !== undefined ? netVal.toFixed(3) : '');
    const slotLabel = num_inputs > 1 ? `<span class="slot-label">Θ${slot + 1}</span>` : '';

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

  const numOpts = Array.from({length:10},(_,i)=>
    `<option value="${i+1}"${num_inputs===i+1?' selected':''}>${i+1}</option>`).join('');
  const tareOpts = Array.from({length:num_inputs+1},(_,i)=>
    `<option value="${i}"${tare_count===i?' selected':''}>${i}</option>`).join('');

  return `
    <div class="item-row${filled ? ' filled' : ''}${num_inputs > 1 ? ' multi-slot' : ''}" id="item-row-${item.row}">
      <div class="item-header-row">
        <div class="item-info">
          <span class="item-name">${h(item.name)}</span>
          ${item.code ? `<span class="item-code">${h(item.code)}</span>` : ''}
        </div>
        <div class="item-cfg-selects">
          <div class="cfg-sel-wrap">
            <span class="cfg-label">Θέσεις</span>
            <select class="cfg-select" onchange="onItemCfgChange(${item.row},'num_inputs',this.value)">${numOpts}</select>
          </div>
          <div class="cfg-sel-wrap">
            <span class="cfg-label">Αποβαρο</span>
            <select class="cfg-select" id="tare-sel-${item.row}" onchange="onItemCfgChange(${item.row},'tare_count',this.value)">${tareOpts}</select>
          </div>
        </div>
      </div>
      <div class="item-slots">
        ${slotsHtml}
        ${totalBar}
      </div>
    </div>`;
}

function switchTab(cat) {
  activeCategory = cat;
  searchQuery = '';
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
  const ic = document.getElementById('items-container');
  if (ic) ic.innerHTML = buildItemsList();
  const tw = document.getElementById('tabs-wrap');
  if (tw) tw.style.display = q ? 'none' : '';
}

function totalCountedText() {
  const n = ITEMS.filter(i => counts[i.row] !== undefined).length;
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

  const cw = containerWeights[itemRow];
  const { num_inputs, tare_count } = getItemCfg(itemRow);
  const isTare = isKilo(item) && slot < tare_count && !!cw;

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
  const total = allSlots.slice(0, num_inputs).reduce((s, v) => s + (v || 0), 0);
  const hasAny = allSlots.slice(0, num_inputs).some(v => v !== undefined);

  if (hasAny) {
    counts[itemRow] = total;
  } else {
    delete counts[itemRow];
  }

  // Update total display
  const totalEl = document.getElementById(`total-${itemRow}`);
  if (totalEl) totalEl.textContent = hasAny ? total.toFixed(3) : '—';

  // Update filled state
  const rowEl = document.getElementById(`item-row-${itemRow}`);
  if (rowEl) rowEl.classList.toggle('filled', hasAny);

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

  const rowEl = document.getElementById(`item-row-${itemRow}`);
  if (rowEl) {
    const tmp = document.createElement('div');
    tmp.innerHTML = buildItemRow(ITEM_MAP[itemRow]);
    rowEl.replaceWith(tmp.firstElementChild);
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
function renderAdminScreen() {
  const kiloItems = ITEMS.filter(isKilo);

  const sections = CATEGORIES.map(cat => {
    const items = kiloItems.filter(i => i.category === cat);
    if (!items.length) return '';
    return `
      <div class="admin-section">
        <h3 class="admin-cat-title">${CAT_LABELS[cat]}</h3>
        ${items.map(item => {
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
                  value="${cwVal}" min="0" step="0.001" inputmode="decimal">
                <span class="unit-label">kg</span>
              </div>
            </div>`;
        }).join('')}
      </div>`;
  }).join('');

  document.getElementById('app').innerHTML = `
    <div class="screen admin-screen">
      <header class="app-header">
        <div class="header-left">
          <button class="back-btn" onclick="renderCountingScreen()">← Πίσω</button>
          <div class="header-title"><strong>⚖️ Βάρη Δοχείων</strong></div>
        </div>
        <div></div>
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
async function exportToExcel() {
  const btn = document.getElementById('export-btn');
  if (btn) { btn.disabled = true; btn.textContent = '⏳ Εξαγωγή...'; }

  try {
    const ws = {};
    const TOTAL_ROWS = 214;

    const setCell = (r, c, value) => {
      if (value === null || value === undefined) return;
      const addr = XLSX.utils.encode_cell({ r, c });
      if (typeof value === 'object' && value.f) {
        ws[addr] = { t: 'n', f: value.f };
      } else if (typeof value === 'number') {
        ws[addr] = { t: 'n', v: value };
      } else {
        ws[addr] = { t: 's', v: String(value) };
      }
    };

    for (let rowNum = 1; rowNum <= TOTAL_ROWS; rowNum++) {
      const r = rowNum - 1; // 0-based index
      const special = SPECIAL_ROWS[rowNum];

      if (special) {
        if (special.empty) continue;
        const d = special.data || [];
        d.forEach((v, c) => setCell(r, c, v));
        if (special.formula) setCell(r, 4, { f: special.formula });
      } else if (ITEM_MAP[rowNum]) {
        const item = ITEM_MAP[rowNum];
        const qty = counts[item.row];
        setCell(r, 0, item.code);
        setCell(r, 1, item.name);
        setCell(r, 2, item.unit_price);
        if (qty !== undefined) setCell(r, 3, qty);
        setCell(r, 4, { f: `C${rowNum}*D${rowNum}` });
        setCell(r, 5, item.unit);
      }
      // else: empty row — no cells written
    }

    ws['!ref'] = XLSX.utils.encode_range({ s: {r:0,c:0}, e: {r:TOTAL_ROWS-1,c:5} });
    ws['!cols'] = [{wch:18},{wch:56},{wch:16},{wch:12},{wch:16},{wch:18}];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Φύλλο1');

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
  if (!initSupabase()) { renderSetupScreen(); return; }

  const savedStore = localStorage.getItem(STORE_KEY);
  const savedDate  = localStorage.getItem(DATE_KEY);
  if (!savedStore || !savedDate) { renderStoreScreen(); return; }

  storeId   = parseInt(savedStore);
  countDate = savedDate;

  document.getElementById('app').innerHTML = `
    <div class="screen loading-screen">
      <div class="spinner"></div>
      <p>Φόρτωση δεδομένων...</p>
    </div>`;

  const ok = await loadData();
  if (ok) { renderCountingScreen(); subscribeToChanges(); }
  else    { renderStoreScreen(); }
}

init();
