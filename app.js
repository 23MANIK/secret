require('dotenv').config();

// Package installations 
const express=require('express');
const app=express();
const mongoose=require('mongoose');
const encrypt=require('mongoose-encryption');
const bodyParser=require('body-parser');
const ejs=require('ejs');
const session=require('express-session');
const passport=require('passport');
const passportLocalMongoose=require('passport-local-mongoose');
const OutlookStrategy = require('passport-outlook').Strategy;
const findOrCreate = require('mongoose-findorcreate');

// starting links and commands
app.use(express.static('public'));
app.set('view engine','ejs');
app.use(bodyParser.urlencoded({extended:true}));
//setup the session
app.use(session({
    secret: "Our Little Secret",
    resave: false,
    saveUninitialized:false
}));

app.use(passport.initialize());
app.use(passport.session());

mongoose.connect("mongodb://localhost:27017/userDB",{useNewUrlParser:true,useUnifiedTopology: true});
mongoose.set("useCreateIndex",true);

const db = mongoose.connection;
db.on('error', console.error.bind(console, 'connection error:'));
db.once('open', function() {
  console.log('we. re connected!') ;
});

//creating database
const userSchema = new mongoose.Schema({
    email:String,
    password:String,
    outlookId: String,
    secret: String
});

userSchema.plugin(passportLocalMongoose);
userSchema.plugin(findOrCreate);

// userSchema.plugin(encrypt, { secret : process.env.SECRET,encryptedFields: ['password'] });

const User=new mongoose.model("User",userSchema);
//serielise and deserielise

passport.use(User.createStrategy());
passport.serializeUser(function(user, done) {
    done(null, user.id);
  });
  
  passport.deserializeUser(function(id, done) {
    User.findById(id, function(err, user) {
      done(err, user);
    });
  });

//strategy

passport.use(new OutlookStrategy({
    clientID: process.env.CLIENT_ID,
    clientSecret: process.env.CLIENT_SECRET,
    callbackURL: 'http://localhost:3000/auth/outlook/secrets',
    passReqToCallback: true
  },
  function(accessToken, refreshToken, profile, done) {
    var user = {
      outlookId: profile.id,
      name: profile.DisplayName,
      email: profile.EmailAddress,
      accessToken:  accessToken
    };
    if (refreshToken)
      user.refreshToken = refreshToken;
    if (profile.MailboxGuid)
      user.mailboxGuid = profile.MailboxGuid;
    if (profile.Alias)
      user.alias = profile.Alias;
    User.findOrCreate(user, function (err, user) {
      return done(err, user);
    });
  }
));
//showing pages
app.get('/',function(req,res){ 
        res.render('home');
});

// app.get('/auth/google',
//     passport.authenticate('google',{scope: ['profile'] })
//     );
app.get('/auth/outlook',
  passport.authenticate('windowslive', {
    scope: [
      'openid',
      'profile',
      'offline_access',
      'https://outlook.office.com/Mail.Read'
    ]
  })
);

//  app.get('/auth/google/secrets', 
//     passport.authenticate('google', { failureRedirect: '/login' }),
//     function(req, res) {
//       // Successful authentication, redirect home.
//       res.redirect('/secrets');
//     });   
app.get('/auth/outlook/secrets', 
  passport.authenticate('windowslive', { failureRedirect: '/login' }),
  function(req, res) {
    // Successful authentication, redirect home.
    res.redirect('/secrets');
  });

app.get('/login',function(req,res){ 
    res.render('login');
});

app.get('/register',function(req,res){ 
    res.render('register');
});

app.get('/secrets',function(req,res){
    User.find({"secret":{$ne:null}},  function(err,foundUsers){
        if(err)console.log(err);
        else
        {
            if(foundUsers)
            {
                res.render("secrets",{userWithSecrets: foundUsers});
            }
        }
    });
});

app.get('/logout',function(req,res){
    req.logOut();
    res.redirect('/');
});

app.get('/submit',function(req,res){
    if(req.isAuthenticated())
    {
        res.render('submit');
    }
    else
    res.redirect('/login');
});

app.post('/submit',function(req,res){
    
    const submittedSecret=req.body.secret;

    console.log(req.user.id);

    User.findById(req.user.id,function(err,foundUser){
        if(err) console.log(err);
        else
        {
            if(foundUser)
            {
                foundUser.secret=submittedSecret;
                foundUser.save(function(){res.redirect('/secrets')})
            }
        }
    })
});

app.post('/register',function(req,res){

    User.register({username: req.body.username},req.body.password,function(err,user){
        if(err){
            console.log(err);
            res.redirect('/register');
        }
        else
        {
                passport.authenticate('local')(req,res,function(){
                    res.redirect('/secrets');
                });
        }
    });

});

app.post('/login',function(req,res){
    
    const user=new User({
        username: req.body.username,
        password: req.body.password
    });
    req.logIn(user,function(err){
        if(err) return console.log(err);
        else
            {
                passport.authenticate("local")(req,res,function(){
                    res.redirect("/secrets");
                });
            }
    })
});


// server running down
app.listen(3000,function(){
    console.log('server is running on 3000');
});