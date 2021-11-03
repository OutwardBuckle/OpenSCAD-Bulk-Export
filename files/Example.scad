message = "X";
textsize = 20;
myfont = "Stencil";

linear_extrude(height = 10){
    text(message, font=myfont,size=textsize);
}