run api :
	cài docker, python (>3.0)
	ùng command prompt trỏ tới thư mục "tag_serice" chạy các lệnh sau:
		docker build -t tag_service .

		docker rm -f tag_service_July18

		docker run -p 3005:5000 --name tag_service_July18 --restart always -t tag_service bash -c "chmod 777 ./scripts/run_service.sh && ./scripts/run_service.sh"

run project:
		cài nodejs
		cài các package trong file package.json của folder "WebApp" bằng lệnh sau:
			npm install + [tên package]
		cài nodemon : npm install nodemon
		trỏ tới folder "WebApp" run : nodemon bin/www