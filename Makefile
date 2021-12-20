ZIP=`which zip`
Productname=TemplateRepo

all:
	./addver.sh; \
	cd src; \
	$(ZIP) -r ../$(Productname).oxt *; \
	cd -; \
	echo -e "\nbuild $(Productname) success..."
