=WENN(ODER($B143=1;$B143=2);
     WENN($H143=1;1;0);
     WENN($K143=1;
          WENN(AllSimilarHaveH1($A143;$B143;$A$143:$A$170;$B$143:$B$170;$H$143:$H$170);1;0);
          0))



=WENN(ODER($B143=1;$B143=2);
     1;
     WENN(HasSimilarAbove($A143;$B143;$A$143:$A$170;$B$143:$B$170);0;1))
		  
