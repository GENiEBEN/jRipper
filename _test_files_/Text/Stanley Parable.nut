num <- 19;		//The overall number
num_ten <- 1;		//The ten place of the number
num_one <- 9;		//Ones place
delay <- 0;		//Randomized tens place delay so that the numbers don't sync up with each other
narration <- false;	//Is the narrator talking?

function Flip()
{
	num+=1;
	num_one+=1;

	if(num==100)	//I sincerely hope that nobody actually waits for the whole thing to flip back to
		num=0;	//zero. Nonetheless, I still add the functionality...but I'm not playtesting it.
	
	if(num_one==10)
	{
		num_ten+=1;
		num_one=0;
		if(num_ten==10)
			num_ten=0;

		delay = (RandomInt(5,10).tofloat())/200;
		
		EntFire("waitroom_sign_ten","Skin",""+num_ten+"",delay);
		EntFire("waitroom_sign_ten","SetAnimation","flip",delay);
	}

	EntFire("waitroom_sign_one","Skin",""+num_one+"",0.00);
	EntFire("waitroom_sign_one","SetAnimation","flip",0.00);

	if(narration)
	{
		if(num==21)
			EntFireByHandle(self,"FireUser1","",0.00,null,null);
		if(num==29)
			EntFireByHandle(self,"FireUser3","",0.00,null,null);
		if(num==32)
			EntFireByHandle(self,"FireUser4","",0.00,null,null);
	}
}

function Reset()
{
	num = 15;
	num_ten = 1;
	num_one = 5;
	EntFire("waitroom_sign_one","Skin",""+(num_one+1)+"",delay);
	EntFire("waitroom_sign_ten","Skin",""+num_ten+"",delay);
	narration = true;
	Flip();
}

function Leave()
{
	narration = false;
	if(num<29)
		EntFireByHandle(self,"FireUser2","",0.00,null,null);
}