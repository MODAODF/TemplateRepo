ZIP=`which zip`
Productname=TemplateRepo

all:
	./addver.sh; \
	cd src; \
	$(ZIP) -r ../$(Productname).oxt *; \
	cd -; \
	mv $(Productname).oxt $(Productname)-`cat version`.oxt
	echo -e "\nbuild $(Productname) success..."
